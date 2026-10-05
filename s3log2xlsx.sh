#!/usr/bin/env bash
# =============================================================================
#  s3log2xlsx.sh - S3 上のログを期間指定で抽出し、S3 のファイルごとにシートを
#                  分けた Excel (.xlsx) を作成する
#
#  動作環境 : RHEL 9 系（bash 5.1 / Python 3.9 の標準ライブラリのみ。追加インストール不要）
#             AWS CLI v2.32.0 以降（aws login に対応したバージョン）
#
#  処理の流れ
#    1. aws login --remote による認証が済んでいるかを確認する（未認証なら警告して終了）
#    2. 現在の権限で対象の S3 を参照できるかを確認する
#         参照できない場合  --switchback warn : スイッチバックを促して終了（既定）
#                           --switchback auto : スイッチバック用シェルを source して続行
#    3. 対象プレフィックス配下の全ログから指定期間の行を抽出し、Excel に出力する
#
#  使い方は「s3log2xlsx.sh --help」を参照してください。
# =============================================================================
set -o nounset -o pipefail

S3X_VERSION="1.0.0"

# -----------------------------------------------------------------------------
# 既定値（環境変数で変更できます。コマンドラインのオプションが最優先です）
# -----------------------------------------------------------------------------
S3X_SWITCHBACK_SCRIPT="${S3LOG2XLSX_SWITCHBACK_SCRIPT:-}"    # スイッチバック用シェルの配置場所
S3X_SWITCHBACK_MODE="${S3LOG2XLSX_SWITCHBACK_MODE:-warn}"    # warn | auto

# 終了コード
S3X_EXIT_ERROR=1
S3X_EXIT_USAGE=2
S3X_EXIT_NOT_LOGGED_IN=10
S3X_EXIT_SWITCHBACK_REQUIRED=11
S3X_EXIT_SWITCHBACK_FAILED=12

# オプション
S3X_S3_URI=""
S3X_BUCKET=""
S3X_PREFIX=""
S3X_FROM=""
S3X_TO="now"
S3X_TZ=""
S3X_LOG_TZ=""
S3X_OUTPUT=""
S3X_KEYWORDS=()
S3X_IGNORE_CASE=0
S3X_REGEX=0
S3X_LOGIN_PROFILE=""
S3X_REGION=""
S3X_SWITCHBACK_ARGS=()
S3X_SKIP_LOGIN_CHECK=0
S3X_INCLUDES=()
S3X_EXCLUDES=()
S3X_MTIME_MARGIN=60
S3X_MAX_FILES=500
S3X_ENCODING=""
S3X_TS_REGEX=""
S3X_TS_FORMAT=""
S3X_UNTIMED="skip"
S3X_WRAP=0
S3X_FORCE=0
S3X_PARALLEL=4
S3X_VERBOSE=0

# 実行中の状態
S3X_WORKDIR=""
S3X_CORE=""
S3X_CORE_ARGS=()
S3X_NOW=""
S3X_LOGIN_PROFILE_RESOLVED=""
S3X_LOGIN_ARN=""
S3X_CURRENT_ARN=""
S3X_PROBE_CATEGORY=""
S3X_PROBE_DETAIL=""
S3X_IN_SWITCHBACK=0

# -----------------------------------------------------------------------------
# メッセージ出力（標準エラー出力へ。端末のときだけ色を付ける）
# -----------------------------------------------------------------------------
if [[ -t 2 && -z ${NO_COLOR:-} ]]; then
    S3X_C_RED=$'\e[31m' S3X_C_YEL=$'\e[33m' S3X_C_GRN=$'\e[32m' S3X_C_CYN=$'\e[36m'
    S3X_C_DIM=$'\e[2m' S3X_C_RST=$'\e[0m'
else
    S3X_C_RED="" S3X_C_YEL="" S3X_C_GRN="" S3X_C_CYN="" S3X_C_DIM="" S3X_C_RST=""
fi

# 1 行目にレベルを付け、2 行目以降は字下げして揃える
s3x__log() {
    local level=$1 color=$2 line first=1
    shift 2
    for line in "$@"; do
        if (( first )); then
            printf '%s[%s]%s %s\n' "$color" "$level" "$S3X_C_RST" "$line" >&2
            first=0
        else
            printf '%*s%s\n' $(( ${#level} + 3 )) '' "$line" >&2
        fi
    done
}
s3x_info()  { s3x__log INFO  "$S3X_C_CYN" "$@"; }
s3x_ok()    { s3x__log OK    "$S3X_C_GRN" "$@"; }
s3x_warn()  { s3x__log WARN  "$S3X_C_YEL" "$@"; }
s3x_error() { s3x__log ERROR "$S3X_C_RED" "$@"; }
s3x_debug() { (( S3X_VERBOSE )) && s3x__log DEBUG "$S3X_C_DIM" "$@"; return 0; }
s3x_die()   { local rc=$1; shift; s3x_error "$@"; exit "$rc"; }

# -----------------------------------------------------------------------------
# ヘルプ
# -----------------------------------------------------------------------------
s3x_usage() {
    local me=${0##*/}
    cat <<EOF
使い方:
  $me -s S3_URI -f 開始日時 [-t 終了日時] -o 出力.xlsx [オプション...]

  S3 の指定プレフィックス配下にあるすべてのログファイルから、指定期間内に出力された
  ログ行を抽出し、S3 のファイルごとにシートを分けた Excel ファイルを作成します。

  実行前に次の 2 点を確認します。
    1. aws login --remote による認証が済んでいるか（未認証なら警告して終了）
    2. 現在の権限で対象の S3 を参照できるか（参照できなければ --switchback に従う）
       ※ 本ツールは CodeCommit を操作しないため、スイッチバックして問題ありません。

必須:
  -s, --s3 URI                 対象の S3 URI（例: s3://my-bucket/app/logs/）
  -f, --from 日時              抽出開始日時
  -o, --output ファイル        出力する Excel ファイル（.xlsx）

期間・タイムゾーン:
  -t, --to 日時                抽出終了日時（既定: 現在時刻）
      --tz TZ                  期間指定と Excel 表示のタイムゾーン（既定: OS の設定）
                               例: Asia/Tokyo  UTC  +09:00
      --log-tz TZ              タイムゾーン表記のないログ日時の解釈（既定: --tz と同じ）

  日時の書式例:
      "2026-10-04 09:00"   "2026/10/04 09:00:30"   "2026-10-04T09:00:00+09:00"
      "2026-10-04"（日付のみ）  "09:00"（本日）  -2h / -30m / -1d（現在から遡る）
      now / today / yesterday
  終了日時は指定した精度の終わりまでを含みます。
      -t "2026-10-04 10:00" → 10:00:59.999 まで / -t 2026-10-04 → 23:59:59.999 まで

強調表示:
  -k, --keyword 文言           強調表示する文言（複数指定可。指定した順に色を割り当てます）
  -i, --ignore-case            強調文言の大文字・小文字を区別しない
      --regex                  強調文言を正規表現として扱う

AWS 認証・権限:
      --login-profile 名前     aws login に使用したプロファイル
                               （既定: AWS_PROFILE → default の順に login_session を持つものを探索）
      --region リージョン      AWS リージョン
      --switchback モード      S3 を参照する権限がないときの動作
                                 warn : スイッチバックを促すメッセージを表示して終了（既定）
                                 auto : スイッチバック用シェルを source して処理を続ける
      --switchback-script パス スイッチバック用シェルの配置場所
      --switchback-arg 引数    スイッチバック用シェルに渡す引数（複数指定可）
      --skip-login-check       aws login の認証チェックを省略する

対象ファイル:
      --include パターン       対象とするオブジェクトキー（ワイルドカード。複数指定可）例: '*.log*'
      --exclude パターン       除外するオブジェクトキー（ワイルドカード。複数指定可）
      --mtime-margin 分        最終更新日時による事前の絞り込みの余裕（既定: 60）
                               最終更新日時が「開始日時 − 余裕」より前のファイルは読み込みません。
                               -1 を指定すると絞り込みません。
      --max-files 件数         読み込むファイル数の上限（既定: 500）
      --encoding 文字コード    文字コードを固定する（既定: UTF-8 / EUC-JP / Shift_JIS を自動判定）
      --parallel 数            同時ダウンロード数（既定: 4）

ログ日時の検出:
  ISO 8601・Apache/Nginx・syslog・Tomcat・JSON（time や timestamp）などの形式を自動判定します。
  日時を含まない行（スタックトレースなど）は、直前のログ行と同じ日時として扱います。
      --ts-regex 正規表現      日時を抜き出す正規表現（最初のグループ、または (?P<ts>...)）
      --ts-format 書式         --ts-regex で抜き出した文字列の書式
                               （strptime 形式 / epoch / epoch_ms / auto。既定: auto）
      --untimed 動作           日時を検出できないファイルの扱い
                                 skip  : 出力しない（既定）
                                 all   : 全行を出力する
                                 mtime : 最終更新日時が期間内なら全行を出力する

Excel:
      --wrap                   ログ列を折り返して表示する（既定: 折り返さない）
  -F, --force                  出力ファイルが既にある場合に上書きする

その他:
  -v, --verbose                詳細なログを表示する
  -h, --help                   このヘルプを表示する
      --version                バージョンを表示する

環境変数:
  S3LOG2XLSX_SWITCHBACK_SCRIPT   --switchback-script の既定値
  S3LOG2XLSX_SWITCHBACK_MODE     --switchback の既定値（warn / auto）

終了コード:
   0  正常終了
   1  エラー
   2  引数の誤り
   3  条件に合うファイルがない（Excel は作成しません）
   4  Excel は作成したが、一部のファイルを読み込めなかった
  10  aws login が未認証、または有効期限切れ
  11  権限不足のためスイッチバックが必要（--switchback warn のとき）
  12  スイッチバック後も S3 を参照できない

例:
  $me -s s3://app-logs/web/ -f "2026-10-04 09:00" -t "2026-10-04 10:00" \\
      -o web_0900-1000.xlsx -k ERROR -k WARN -k Timeout

  $me -s s3://app-logs/batch/ -f -3h -o batch.xlsx -k Exception \\
      --switchback auto --switchback-script /opt/team/bin/switchback.sh
EOF
}

# -----------------------------------------------------------------------------
# 引数の解析
# -----------------------------------------------------------------------------
# 値を取るオプションの値があるか（呼び出し元の opt を参照する）
s3x__need_value() {
    (( $# > 0 )) || s3x_die "$S3X_EXIT_USAGE" "オプション $opt には値が必要です（--help を参照）"
}
# 値を取らないオプションに --opt=値 の形で値が付いていないか（呼び出し元の inline を参照する）
s3x__no_value() {
    (( ! inline )) || s3x_die "$S3X_EXIT_USAGE" "オプション $opt は値を取りません（--help を参照）"
}

s3x_parse_args() {
    local opt inline
    while (( $# > 0 )); do
        opt=$1
        shift
        inline=0
        if [[ $opt == --?*=* ]]; then
            # --opt=値 は「--opt 値」に置き換えて同じように扱う
            set -- "${opt#*=}" "$@"
            opt=${opt%%=*}
            inline=1
        fi
        case $opt in
            -s|--s3)              s3x__need_value "$@"; S3X_S3_URI=$1; shift ;;
            -f|--from)            s3x__need_value "$@"; S3X_FROM=$1; shift ;;
            -t|--to)              s3x__need_value "$@"; S3X_TO=$1; shift ;;
            -o|--output)          s3x__need_value "$@"; S3X_OUTPUT=$1; shift ;;
            --tz)                 s3x__need_value "$@"; S3X_TZ=$1; shift ;;
            --log-tz)             s3x__need_value "$@"; S3X_LOG_TZ=$1; shift ;;
            -k|--keyword)         s3x__need_value "$@"; S3X_KEYWORDS+=("$1"); shift ;;
            -i|--ignore-case)     s3x__no_value; S3X_IGNORE_CASE=1 ;;
            --regex)              s3x__no_value; S3X_REGEX=1 ;;
            --login-profile)      s3x__need_value "$@"; S3X_LOGIN_PROFILE=$1; shift ;;
            --region)             s3x__need_value "$@"; S3X_REGION=$1; shift ;;
            --switchback)         s3x__need_value "$@"; S3X_SWITCHBACK_MODE=$1; shift ;;
            --switchback-script)  s3x__need_value "$@"; S3X_SWITCHBACK_SCRIPT=$1; shift ;;
            --switchback-arg)     s3x__need_value "$@"; S3X_SWITCHBACK_ARGS+=("$1"); shift ;;
            --skip-login-check)   s3x__no_value; S3X_SKIP_LOGIN_CHECK=1 ;;
            --include)            s3x__need_value "$@"; S3X_INCLUDES+=("$1"); shift ;;
            --exclude)            s3x__need_value "$@"; S3X_EXCLUDES+=("$1"); shift ;;
            --mtime-margin)       s3x__need_value "$@"; S3X_MTIME_MARGIN=$1; shift ;;
            --max-files)          s3x__need_value "$@"; S3X_MAX_FILES=$1; shift ;;
            --encoding)           s3x__need_value "$@"; S3X_ENCODING=$1; shift ;;
            --parallel)           s3x__need_value "$@"; S3X_PARALLEL=$1; shift ;;
            --ts-regex)           s3x__need_value "$@"; S3X_TS_REGEX=$1; shift ;;
            --ts-format)          s3x__need_value "$@"; S3X_TS_FORMAT=$1; shift ;;
            --untimed)            s3x__need_value "$@"; S3X_UNTIMED=$1; shift ;;
            --wrap)               s3x__no_value; S3X_WRAP=1 ;;
            -F|--force)           s3x__no_value; S3X_FORCE=1 ;;
            -v|--verbose)         s3x__no_value; S3X_VERBOSE=1 ;;
            -h|--help)            s3x_usage; exit 0 ;;
            --version)            printf 's3log2xlsx %s\n' "$S3X_VERSION"; exit 0 ;;
            *)                    s3x_die "$S3X_EXIT_USAGE" "不明な引数です: $opt（--help を参照）" ;;
        esac
    done
}

# ~/ で始まるパスと相対パスを絶対パスにする（source が PATH を探さないように）
s3x_abs_path() {
    local p=$1
    [[ $p == \~/* ]] && p="$HOME/${p#\~/}"
    [[ $p == /* ]] || p="$PWD/$p"
    printf '%s' "$p"
}

s3x_validate_args() {
    [[ -n $S3X_S3_URI ]] || s3x_die "$S3X_EXIT_USAGE" "-s/--s3 で対象の S3 URI を指定してください（--help を参照）"
    [[ -n $S3X_FROM ]]   || s3x_die "$S3X_EXIT_USAGE" "-f/--from で抽出開始日時を指定してください（--help を参照）"
    [[ -n $S3X_OUTPUT ]] || s3x_die "$S3X_EXIT_USAGE" "-o/--output で出力する Excel ファイルを指定してください（--help を参照）"

    local uri=$S3X_S3_URI
    [[ ${uri,,} == s3://* ]] && uri=${uri:5}
    S3X_BUCKET=${uri%%/*}
    if [[ $uri == */* ]]; then S3X_PREFIX=${uri#*/}; else S3X_PREFIX=""; fi
    [[ $S3X_BUCKET =~ ^[A-Za-z0-9][A-Za-z0-9._-]{1,254}$ ]] ||
        s3x_die "$S3X_EXIT_USAGE" "S3 URI のバケット名が正しくありません: $S3X_S3_URI（例: s3://my-bucket/logs/）"

    case $S3X_SWITCHBACK_MODE in
        warn|auto) ;;
        *) s3x_die "$S3X_EXIT_USAGE" "--switchback には warn または auto を指定してください: $S3X_SWITCHBACK_MODE" ;;
    esac
    if [[ -n $S3X_SWITCHBACK_SCRIPT ]]; then
        S3X_SWITCHBACK_SCRIPT=$(s3x_abs_path "$S3X_SWITCHBACK_SCRIPT")
    fi
    if [[ $S3X_SWITCHBACK_MODE == auto ]]; then
        [[ -n $S3X_SWITCHBACK_SCRIPT ]] ||
            s3x_die "$S3X_EXIT_USAGE" "--switchback auto には、スイッチバック用シェルの指定が必要です。" \
                "--switchback-script または環境変数 S3LOG2XLSX_SWITCHBACK_SCRIPT で配置場所を指定してください。"
        [[ -f $S3X_SWITCHBACK_SCRIPT && -r $S3X_SWITCHBACK_SCRIPT ]] ||
            s3x_die "$S3X_EXIT_USAGE" "スイッチバック用シェルが見つからないか、読み込めません: $S3X_SWITCHBACK_SCRIPT"
    fi

    case $S3X_UNTIMED in
        skip|all|mtime) ;;
        *) s3x_die "$S3X_EXIT_USAGE" "--untimed には skip / all / mtime のいずれかを指定してください: $S3X_UNTIMED" ;;
    esac
    [[ $S3X_MTIME_MARGIN =~ ^(-1|[0-9]+)$ ]] ||
        s3x_die "$S3X_EXIT_USAGE" "--mtime-margin には 0 以上の分数、または -1 を指定してください: $S3X_MTIME_MARGIN"
    [[ $S3X_MAX_FILES =~ ^[1-9][0-9]*$ ]] ||
        s3x_die "$S3X_EXIT_USAGE" "--max-files には 1 以上の整数を指定してください: $S3X_MAX_FILES"
    if [[ ! $S3X_PARALLEL =~ ^[1-9][0-9]?$ ]] || (( S3X_PARALLEL > 16 )); then
        s3x_die "$S3X_EXIT_USAGE" "--parallel には 1〜16 の整数を指定してください: $S3X_PARALLEL"
    fi
    return 0
}

# -----------------------------------------------------------------------------
# 前提コマンドの確認
# -----------------------------------------------------------------------------
s3x_check_prereqs() {
    command -v aws >/dev/null 2>&1 ||
        s3x_die "$S3X_EXIT_ERROR" "aws コマンドが見つかりません。AWS CLI v2 をインストールしてください。"
    command -v python3 >/dev/null 2>&1 ||
        s3x_die "$S3X_EXIT_ERROR" "python3 が見つかりません。"
    python3 -c 'import sys; sys.exit(0 if sys.version_info >= (3, 9) else 1)' 2>/dev/null ||
        s3x_die "$S3X_EXIT_ERROR" "Python 3.9 以降が必要です（現在: $(python3 --version 2>&1)）。"

    if (( ! S3X_SKIP_LOGIN_CHECK )); then
        local ver
        ver=$(aws --version 2>&1 | head -n 1)
        if [[ $ver =~ aws-cli/([0-9]+)\.([0-9]+)\.[0-9]+ ]]; then
            if (( BASH_REMATCH[1] < 2 || (BASH_REMATCH[1] == 2 && BASH_REMATCH[2] < 32) )); then
                s3x_warn "AWS CLI が古いため aws login を利用できない可能性があります（現在: ${ver%% *}、必要: aws-cli/2.32.0 以降）。"
            fi
        fi
    fi
}

# -----------------------------------------------------------------------------
# aws コマンドの呼び出し
# -----------------------------------------------------------------------------
s3x_aws() {
    aws "$@" ${S3X_REGION:+--region "$S3X_REGION"}
}

# aws login の認証状態を確かめるための呼び出し。スイッチロールなどで環境変数に
# 設定された認証情報の影響を受けないように、それらを外して実行する。
s3x_aws_isolated() {
    env -u AWS_ACCESS_KEY_ID -u AWS_SECRET_ACCESS_KEY -u AWS_SESSION_TOKEN -u AWS_SECURITY_TOKEN \
        -u AWS_CREDENTIAL_EXPIRATION -u AWS_PROFILE -u AWS_DEFAULT_PROFILE \
        -u AWS_ROLE_ARN -u AWS_ROLE_SESSION_NAME -u AWS_WEB_IDENTITY_TOKEN_FILE \
        aws "$@" ${S3X_REGION:+--region "$S3X_REGION"}
}

# aws のエラー出力から要点（エラーコードとメッセージ）を 1 行で取り出す
s3x_aws_error_line() {
    local line last=""
    while IFS= read -r line; do
        line=${line#aws: \[ERROR\]: }
        [[ -n ${line//[[:space:]]/} ]] && last=$line
    done <<<"$1"
    printf '%s' "${last:-不明なエラー}"
}

# aws のエラーを分類する
#   denied   : 権限・認証情報の問題（スイッチバックで解消する可能性がある）
#   nobucket : バケットが存在しない
#   network  : 通信エラー
#   region   : リージョンの指定誤り
#   other    : その他
s3x_classify_aws_error() {
    case $1 in
        *NoSuchBucket*) echo nobucket ;;
        *"Could not connect"*|*"Connect timeout"*|*"Read timeout"*|*"Connection was closed"*|*EndpointConnectionError*)
            echo network ;;
        *PermanentRedirect*|*AuthorizationHeaderMalformed*|*IllegalLocationConstraint*)
            echo region ;;
        *AccessDenied*|*AllAccessDisabled*|*"(403)"*|*Forbidden*|*ExpiredToken*|*InvalidAccessKeyId*|\
        *InvalidToken*|*SignatureDoesNotMatch*|*InvalidClientTokenId*|*UnrecognizedClient*|*NoCredentials*|\
        *"Unable to locate credentials"*|*"could not be found"*|*"login session"*|*reauthenticate*|\
        *"Error when retrieving"*|*"Token has expired"*|*"token has expired"*)
            echo denied ;;
        *) echo other ;;
    esac
}

# -----------------------------------------------------------------------------
# 1. aws login --remote による認証の確認
# -----------------------------------------------------------------------------
# login_session を持つプロファイルを探す（source_profile をたどる）
s3x_resolve_login_profile() {
    local -a candidates=()
    if [[ -n $S3X_LOGIN_PROFILE ]]; then
        candidates=("$S3X_LOGIN_PROFILE")
    else
        [[ -n ${AWS_PROFILE:-} ]] && candidates+=("$AWS_PROFILE")
        [[ ${AWS_PROFILE:-} == default ]] || candidates+=(default)
    fi
    local cand p src session depth
    for cand in "${candidates[@]}"; do
        p=$cand
        for (( depth = 0; depth < 5; depth++ )); do
            session=$(s3x_aws_isolated configure get login_session --profile "$p" 2>/dev/null) || session=""
            if [[ -n $session ]]; then
                S3X_LOGIN_PROFILE_RESOLVED=$p
                s3x_debug "aws login のプロファイル: $p（login_session = $session）"
                return 0
            fi
            src=$(s3x_aws_isolated configure get source_profile --profile "$p" 2>/dev/null) || src=""
            [[ -n $src && $src != "$p" ]] || break
            p=$src
        done
    done
    return 1
}

s3x_check_login() {
    local hint="aws login --remote"
    [[ -n $S3X_LOGIN_PROFILE && $S3X_LOGIN_PROFILE != default ]] && hint+=" --profile $S3X_LOGIN_PROFILE"

    if ! s3x_resolve_login_profile; then
        local where="default"
        [[ -n $S3X_LOGIN_PROFILE ]] && where=$S3X_LOGIN_PROFILE
        [[ -z $S3X_LOGIN_PROFILE && -n ${AWS_PROFILE:-} && ${AWS_PROFILE} != default ]] && where="${AWS_PROFILE}、default"
        s3x_error "AWS に未認証です。aws login --remote による認証が行われていません。" \
            "（プロファイル ${where} に login_session の設定がありません）" \
            "次のコマンドで認証してから、もう一度実行してください。" \
            "    $hint"
        exit "$S3X_EXIT_NOT_LOGGED_IN"
    fi

    local p=$S3X_LOGIN_PROFILE_RESOLVED out rc
    [[ $p != default ]] && hint="aws login --remote --profile $p"
    out=$(s3x_aws_isolated sts get-caller-identity --profile "$p" --query Arn --output text 2>&1)
    rc=$?
    if (( rc != 0 )); then
        local detail
        detail=$(s3x_aws_error_line "$out")
        if [[ $(s3x_classify_aws_error "$out") == network ]]; then
            s3x_die "$S3X_EXIT_ERROR" "aws login の認証状態を確認できませんでした（AWS に接続できません）。" \
                "詳細: $detail"
        fi
        s3x_error "AWS に未認証、または aws login の認証の有効期限が切れています（プロファイル: $p）。" \
            "次のコマンドで認証してから、もう一度実行してください。" \
            "    $hint" \
            "詳細: $detail"
        exit "$S3X_EXIT_NOT_LOGGED_IN"
    fi
    S3X_LOGIN_ARN=${out//[[:space:]]/}
    s3x_ok "aws login 認証済み: $S3X_LOGIN_ARN（プロファイル: $p）"

    # 認証情報の指定が何もない状態なら、ログインしたプロファイルで S3 を操作する
    if [[ -z ${AWS_PROFILE:-} && -z ${AWS_ACCESS_KEY_ID:-} && $p != default ]]; then
        export AWS_PROFILE=$p
        s3x_debug "AWS_PROFILE=$p を設定しました"
    fi
}

# -----------------------------------------------------------------------------
# 2. 対象の S3 を参照する権限の確認とスイッチバック
# -----------------------------------------------------------------------------
# 現在の認証情報の ID（表示用）
s3x_current_identity() {
    local out
    S3X_CURRENT_ARN=""
    if out=$(s3x_aws sts get-caller-identity --query Arn --output text 2>&1); then
        S3X_CURRENT_ARN=${out//[[:space:]]/}
    else
        s3x_debug "現在の ID を取得できませんでした: $(s3x_aws_error_line "$out")"
    fi
    return 0
}

# 一覧の取得（s3:ListBucket）と先頭オブジェクトの読み取り（s3:GetObject）を試す。
# 判定は実際の S3 の応答で行う（STS に届かない環境でも S3 を参照できる場合があるため）。
s3x_probe_s3_access() {
    local out key size
    S3X_PROBE_CATEGORY="" S3X_PROBE_DETAIL=""
    local -a prefix_opt=()
    [[ -n $S3X_PREFIX ]] && prefix_opt=("--prefix=$S3X_PREFIX")
    if ! out=$(s3x_aws s3api list-objects-v2 "--bucket=$S3X_BUCKET" "${prefix_opt[@]}" \
            --max-items 1 --page-size 1 --query 'Contents[0].[Key,Size]' --output text 2>&1); then
        S3X_PROBE_CATEGORY=$(s3x_classify_aws_error "$out")
        S3X_PROBE_DETAIL=$(s3x_aws_error_line "$out")
        return 1
    fi
    out=${out%$'\n'}
    [[ -z $out || $out == None ]] && return 0      # オブジェクトなし（一覧の権限はある）
    key=${out%$'\t'*}
    size=${out##*$'\t'}
    local -a read_cmd=(s3api head-object "--bucket=$S3X_BUCKET" "--key=$key")
    if [[ $size =~ ^[0-9]+$ ]] && (( size > 0 )); then
        # 1 バイトだけ取得する（KMS で暗号化されたオブジェクトの復号権限も確認できる）
        read_cmd=(s3api get-object "--bucket=$S3X_BUCKET" "--key=$key" --range bytes=0-0 /dev/null)
    fi
    if ! out=$(s3x_aws "${read_cmd[@]}" --query ContentLength --output text 2>&1); then
        S3X_PROBE_CATEGORY=$(s3x_classify_aws_error "$out")
        S3X_PROBE_DETAIL="$(s3x_aws_error_line "$out")（$key）"
        return 1
    fi
    return 0
}

s3x_report_switchback_required() {
    local -a msg=(
        "現在の権限では対象の S3（s3://$S3X_BUCKET/$S3X_PREFIX）を参照できません。"
        "  理由        : $S3X_PROBE_DETAIL"
        "  現在の ID   : ${S3X_CURRENT_ARN:-（取得できませんでした）}"
    )
    if [[ -n $S3X_LOGIN_ARN && $S3X_LOGIN_ARN != "$S3X_CURRENT_ARN" ]]; then
        msg+=("  ログイン ID : $S3X_LOGIN_ARN")
    fi
    msg+=("本ツールは CodeCommit を操作しないため、スイッチバックしてから再実行してください。")
    if [[ -n $S3X_SWITCHBACK_SCRIPT ]]; then
        local cmd a
        cmd="source $(printf '%q' "$S3X_SWITCHBACK_SCRIPT")"
        for a in "${S3X_SWITCHBACK_ARGS[@]}"; do cmd+=" $(printf '%q' "$a")"; done
        msg+=("    $cmd")
        [[ -f $S3X_SWITCHBACK_SCRIPT ]] || msg+=("  ※ 上記のシェルが見つかりません。配置場所を確認してください。")
        msg+=("自動でスイッチバックさせる場合は --switchback auto を指定してください。")
    else
        msg+=("スイッチバック用のシェルの場所を --switchback-script（または環境変数 S3LOG2XLSX_SWITCHBACK_SCRIPT）"
              "で指定すると、--switchback auto で自動的にスイッチバックできます。")
    fi
    s3x_warn "${msg[@]}"
}

# スイッチバック用シェルを source する。呼び出し先がシェルの設定（set -e / set -u /
# トラップ / IFS / カレントディレクトリなど）を変えても本ツールが影響を受けないよう、
# 実行前の状態を保存して元に戻す。
s3x_run_switchback() {
    local script=$S3X_SWITCHBACK_SCRIPT rc v cmd a
    local saved_set saved_shopt saved_traps saved_pwd=$PWD saved_umask saved_ifs=$IFS
    cmd="source $(printf '%q' "$script")"
    for a in "${S3X_SWITCHBACK_ARGS[@]}"; do cmd+=" $(printf '%q' "$a")"; done
    s3x_info "スイッチバックを自動で実行します: $cmd"

    saved_set=$(set +o)
    saved_shopt=$(shopt -p)
    saved_traps=$(trap -p)
    saved_umask=$(umask)
    set +o errexit +o nounset
    S3X_IN_SWITCHBACK=1
    # shellcheck source=/dev/null
    source "$script" "${S3X_SWITCHBACK_ARGS[@]}"
    rc=$?
    S3X_IN_SWITCHBACK=0
    trap - DEBUG RETURN ERR
    eval "$saved_traps"
    eval "$saved_shopt" 2>/dev/null
    eval "$saved_set"
    IFS=$saved_ifs
    umask "$saved_umask"
    cd "$saved_pwd" 2>/dev/null || true

    # 呼び出し先が export せずに設定した AWS_ で始まる変数も、aws コマンドに渡るようにする
    while IFS= read -r v; do
        [[ $v == AWS_* ]] && export "${v?}"
    done < <(compgen -v)

    if (( rc != 0 )); then
        s3x_warn "スイッチバック用シェルが終了コード $rc を返しました。権限をもう一度確認します。"
    fi
}

s3x_ensure_s3_access() {
    s3x_current_identity
    if s3x_probe_s3_access; then
        s3x_ok "S3 を参照できることを確認しました（ID: ${S3X_CURRENT_ARN:-不明}）"
        return 0
    fi

    case $S3X_PROBE_CATEGORY in
        denied) ;;
        nobucket)
            s3x_die "$S3X_EXIT_ERROR" "バケットが存在しません: s3://$S3X_BUCKET/（$S3X_PROBE_DETAIL）" ;;
        network)
            s3x_die "$S3X_EXIT_ERROR" "S3 に接続できません。ネットワークやプロキシの設定を確認してください。" \
                "詳細: $S3X_PROBE_DETAIL" ;;
        region)
            s3x_die "$S3X_EXIT_ERROR" "バケットのリージョンが一致しません。--region でバケットのリージョンを指定してください。" \
                "詳細: $S3X_PROBE_DETAIL" ;;
        *)
            s3x_die "$S3X_EXIT_ERROR" "S3 の確認中にエラーが発生しました。" "詳細: $S3X_PROBE_DETAIL" ;;
    esac

    if [[ $S3X_SWITCHBACK_MODE != auto ]]; then
        s3x_report_switchback_required
        exit "$S3X_EXIT_SWITCHBACK_REQUIRED"
    fi

    s3x_warn "現在の権限では対象の S3（s3://$S3X_BUCKET/$S3X_PREFIX）を参照できないため、スイッチバックします。" \
        "  理由      : $S3X_PROBE_DETAIL" \
        "  現在の ID : ${S3X_CURRENT_ARN:-（取得できませんでした）}"
    s3x_run_switchback
    s3x_current_identity
    if s3x_probe_s3_access; then
        s3x_ok "スイッチバック後、S3 を参照できることを確認しました（ID: ${S3X_CURRENT_ARN:-不明}）" \
            "※ スイッチバックは本ツールの処理中だけ有効です。実行元のシェルの状態は変わりません。"
        return 0
    fi
    s3x_error "スイッチバック後も対象の S3 を参照できません。" \
        "  理由      : $S3X_PROBE_DETAIL" \
        "  現在の ID : ${S3X_CURRENT_ARN:-（取得できませんでした）}" \
        "スイッチバック用シェル（$S3X_SWITCHBACK_SCRIPT）の内容と、対象バケットの権限を確認してください。"
    exit "$S3X_EXIT_SWITCHBACK_FAILED"
}

# -----------------------------------------------------------------------------
# 3. ログの抽出と Excel の作成（Python 部分に引き渡す）
# -----------------------------------------------------------------------------
s3x_build_core_args() {
    local x
    S3X_CORE_ARGS=(
        "--bucket=$S3X_BUCKET" "--prefix=$S3X_PREFIX"
        "--from=$S3X_FROM" "--to=$S3X_TO" "--now=$S3X_NOW"
        "--tz=$S3X_TZ" "--log-tz=$S3X_LOG_TZ"
        "--output=$S3X_OUTPUT"
        "--mtime-margin=$S3X_MTIME_MARGIN" "--max-files=$S3X_MAX_FILES" "--parallel=$S3X_PARALLEL"
        "--encoding=$S3X_ENCODING" "--ts-regex=$S3X_TS_REGEX" "--ts-format=$S3X_TS_FORMAT"
        "--untimed=$S3X_UNTIMED" "--region=$S3X_REGION"
        "--workdir=$S3X_WORKDIR" "--tool-version=$S3X_VERSION"
    )
    for x in "${S3X_KEYWORDS[@]}"; do S3X_CORE_ARGS+=("--keyword=$x"); done
    for x in "${S3X_INCLUDES[@]}"; do S3X_CORE_ARGS+=("--include=$x"); done
    for x in "${S3X_EXCLUDES[@]}"; do S3X_CORE_ARGS+=("--exclude=$x"); done
    (( S3X_IGNORE_CASE )) && S3X_CORE_ARGS+=(--ignore-case)
    (( S3X_REGEX ))       && S3X_CORE_ARGS+=(--regex)
    (( S3X_WRAP ))        && S3X_CORE_ARGS+=(--wrap)
    (( S3X_FORCE ))       && S3X_CORE_ARGS+=(--force)
    (( S3X_VERBOSE ))     && S3X_CORE_ARGS+=(--verbose)
    return 0
}

s3x_on_exit() {
    local rc=$?
    if (( S3X_IN_SWITCHBACK )); then
        s3x_error "スイッチバック用シェル（$S3X_SWITCHBACK_SCRIPT）の中で exit が実行されたため、処理を中断しました（終了コード: $rc）。" \
            "source で呼び出されることを想定したシェルか確認してください。"
    fi
    if [[ -n $S3X_WORKDIR && -d $S3X_WORKDIR ]]; then
        rm -rf -- "$S3X_WORKDIR"
    fi
    return "$rc"
}

s3x_main() {
    s3x_parse_args "$@"
    s3x_validate_args
    s3x_check_prereqs

    export AWS_PAGER="" AWS_CLI_AUTO_PROMPT=off
    S3X_NOW=$(date +%s.%N)

    trap s3x_on_exit EXIT
    trap 'exit 130' INT
    trap 'exit 143' TERM
    S3X_WORKDIR=$(mktemp -d "${TMPDIR:-/tmp}/s3log2xlsx.XXXXXXXX") ||
        s3x_die "$S3X_EXIT_ERROR" "作業用ディレクトリを作成できません（TMPDIR=${TMPDIR:-/tmp}）。"
    S3X_CORE="$S3X_WORKDIR/s3log2xlsx_core.py"
    s3x_write_core >"$S3X_CORE" || s3x_die "$S3X_EXIT_ERROR" "作業用ファイルを作成できません: $S3X_CORE"
    s3x_build_core_args

    s3x_info "s3log2xlsx $S3X_VERSION"
    python3 "$S3X_CORE" check "${S3X_CORE_ARGS[@]}" || exit $?

    if (( S3X_SKIP_LOGIN_CHECK )); then
        s3x_warn "aws login の認証チェックを省略します（--skip-login-check）。"
    else
        s3x_check_login
    fi
    s3x_ensure_s3_access

    python3 "$S3X_CORE" run "${S3X_CORE_ARGS[@]}" "--identity=$S3X_CURRENT_ARN"
}

# -----------------------------------------------------------------------------
# Python 部分（Python 3.9 の標準ライブラリのみを使用）
# -----------------------------------------------------------------------------
s3x_write_core() {
    cat <<'S3X_PY_EOF'
# -*- coding: utf-8 -*-
# =============================================================================
#  s3log2xlsx core（s3log2xlsx.sh から呼び出される内部処理）
#    - aws CLI による S3 のオブジェクト一覧の取得とダウンロード
#    - 圧縮形式（gzip / bzip2 / xz）、文字コード（UTF-8 / EUC-JP / Shift_JIS）、
#      ログの日時形式の自動判定
#    - 期間内の行の抽出と、強調文言の色付け
#    - Excel (.xlsx) の作成（Office Open XML を直接書き出すため外部ライブラリは不要）
# =============================================================================
import argparse
import bz2
import codecs
import collections
import concurrent.futures
import datetime as dt
import fnmatch
import gzip
import io
import itertools
import json
import lzma
import os
import re
import shlex
import shutil
import signal
import subprocess
import sys
import tempfile
import time
import unicodedata
import zipfile
import zlib

try:
    from zoneinfo import ZoneInfo
except ImportError:  # Python 3.8 以前
    ZoneInfo = None

EXIT_OK, EXIT_ERROR, EXIT_USAGE, EXIT_NO_TARGET, EXIT_PARTIAL = 0, 1, 2, 3, 4

UTC = dt.timezone.utc
FONT_NAME = 'Meiryo UI'
EXCEL_MAX_ROWS = 1048576
CELL_LIMIT = 32767          # 1 セルに入る最大文字数（UTF-16 単位）
HEADER_ROWS = 4             # データシートの見出し行数（タイトル・情報 2 行・表の見出し）
ROW_HT = 16.5               # ログの行の高さ（ポイント）
SAMPLE_LINES = 5000         # 日時形式の判定に使う先頭の行数
WIDTH_MAX = 400             # 列幅計算用ヒストグラムの上限
DETECT_LIMIT = 32 << 20     # 文字コードの判定で読む最大バイト数
SUMMARY_SHEET = 'サマリー'

# 配色（落ち着いたネイビーを基調にし、強調文言は彩度の高い色で区別する）
C_NAVY = '1F3A5F'
C_NAVY_LINE = '3A5A80'
C_TEXT = '24292F'
C_SUB = '57606A'
C_MUTED = '8C959F'
C_FAINT = 'B1BAC4'
C_BAND = 'F6F8FA'
C_GRID = 'D8DEE4'
C_LINK = '0969DA'
C_WHITE = 'FFFFFF'
KW_PALETTE = (             # (文字色, 行の背景色)
    ('CF222E', 'FFEBE9'),  # 赤
    ('9A6700', 'FFF8C5'),  # 黄
    ('0969DA', 'DDF4FF'),  # 青
    ('1A7F37', 'DAFBE1'),  # 緑
    ('8250DF', 'FBEFFF'),  # 紫
    ('BC4C00', 'FFF1E5'),  # 橙
    ('BF3989', 'FFEFF7'),  # 桃
    ('1B7C83', 'DFF6F5'),  # 青緑
)
STATUS_COLOR = {'ok': '1A7F37', 'muted': C_MUTED, 'warn': '9A6700', 'error': 'CF222E'}

FMT_TS = 'yyyy/mm/dd hh:mm:ss.000'
FMT_DT = 'yyyy/mm/dd hh:mm:ss'
FMT_INT = '#,##0'
FMT_SIZE = '[<1000]0" B";[<1000000]0.0," KB";#,##0.0,," MB"'


class UsageError(Exception):
    """引数の誤り（終了コード 2）"""


class FatalError(Exception):
    """処理を続けられないエラー（終了コード 1）"""


class Terminated(BaseException):
    """SIGTERM を受け取った"""


# =============================================================================
# メッセージ出力
# =============================================================================
def disp_width(s):
    """端末での表示幅（全角を 2 とする）"""
    return sum(2 if unicodedata.east_asian_width(ch) in 'WFA' else 1 for ch in s)


class Console:
    def __init__(self):
        self.color = sys.stderr.isatty() and not os.environ.get('NO_COLOR')
        self.verbose = False
        try:
            sys.stderr.reconfigure(errors='replace')
        except (AttributeError, ValueError):
            pass

    def _emit(self, label, color, lines):
        head = ('\033[%sm[%s]\033[0m ' % (color, label)) if self.color else '[%s] ' % label
        pad = ' ' * (len(label) + 3)
        out = [(head if i == 0 else pad) + line for i, line in enumerate(lines)]
        sys.stderr.write('\n'.join(out) + '\n')
        sys.stderr.flush()

    def info(self, *lines):
        self._emit('INFO', '36', lines)

    def ok(self, *lines):
        self._emit('OK', '32', lines)

    def warn(self, *lines):
        self._emit('WARN', '33', lines)

    def error(self, *lines):
        self._emit('ERROR', '31', lines)

    def debug(self, *lines):
        if self.verbose:
            self._emit('DEBUG', '2', lines)

    def kv(self, label, value):
        self.info('%s%s : %s' % (label, ' ' * max(0, 8 - disp_width(label)), value))

    def progress(self, i, n, text, level='info'):
        head = '[%*d/%d]' % (len(str(n)), i, n)
        color = {'warn': '33', 'error': '31'}.get(level)
        if self.color:
            head = '\033[2m%s\033[0m' % head
            if color:
                text = '\033[%sm%s\033[0m' % (color, text)
        sys.stderr.write('%s %s\n' % (head, text))
        sys.stderr.flush()


def fmt_n(n):
    return format(n, ',')


def human_size(n):
    if n < 1000:
        return '%d B' % n
    if n < 1000000:
        return '%.1f KB' % (n / 1000.0)
    return '%s MB' % format(n / 1000000.0, ',.1f')


# =============================================================================
# タイムゾーンと日時
# =============================================================================
def _zone(name):
    if ZoneInfo is None:
        raise ValueError(name)
    try:
        return ZoneInfo(name)
    except Exception:
        raise ValueError(name)


_TZ_OFFSETS = {}


def tz_offset(s):
    """'Z' / 'UTC' / '+09:00' / '+0900' / '+09' を UTC からの秒数に変換する（不正なら None）"""
    v = _TZ_OFFSETS.get(s, False)
    if v is not False:
        return v
    u = s.upper()
    if u in ('Z', 'UTC', 'GMT'):
        v = 0
    elif u == 'JST':
        v = 9 * 3600
    else:
        body = s[1:].replace(':', '')
        v = None
        if s[:1] in '+-' and body.isdigit() and len(body) in (2, 4):
            h, mi = int(body[:2]), int(body[2:4] or 0)
            if h <= 14 and mi < 60:
                v = (-1 if s[0] == '-' else 1) * (h * 3600 + mi * 60)
    _TZ_OFFSETS[s] = v
    return v


def parse_tz(spec):
    s = (spec or '').strip()
    if not s or s.lower() == 'local':
        return local_tz()
    u = s.upper()
    if u in ('UTC', 'Z', 'GMT', 'ETC/UTC', 'UTC0'):
        return UTC
    if u == 'JST':
        try:
            return _zone('Asia/Tokyo')
        except ValueError:
            return dt.timezone(dt.timedelta(hours=9), 'JST')
    m = re.fullmatch(r'(?:UTC|GMT)?\s*([+-])(\d{1,2})(?::?(\d{2}))?', s, re.I)
    if m:
        h, mi = int(m.group(2)), int(m.group(3) or 0)
        if h > 14 or mi > 59:
            raise ValueError(spec)
        off = dt.timedelta(hours=h, minutes=mi) * (-1 if m.group(1) == '-' else 1)
        return dt.timezone(off) if off else UTC
    return _zone(s)


def local_tz():
    """OS のタイムゾーン（環境変数 TZ → /etc/localtime の順）"""
    name = os.environ.get('TZ', '').strip().lstrip(':')
    if name and name.lower() != 'local':
        try:
            return parse_tz(name)
        except ValueError:
            pass
    try:
        target = os.path.realpath('/etc/localtime')
        i = target.find('/zoneinfo/')
        if i >= 0:
            return _zone(target[i + 10:])
    except (OSError, ValueError):
        pass
    if ZoneInfo is not None:
        try:
            with open('/etc/localtime', 'rb') as f:
                return ZoneInfo.from_file(f, key='localtime')
        except Exception:
            pass
    return dt.timezone(dt.datetime.now().astimezone().utcoffset() or dt.timedelta(0))


def _offset_text(seconds):
    seconds = int(seconds)
    sign = '+' if seconds >= 0 else '-'
    seconds = abs(seconds)
    return 'UTC%s%02d:%02d' % (sign, seconds // 3600, seconds % 3600 // 60)


def tz_long(tz, ts):
    """例: 'Asia/Tokyo, UTC+09:00'"""
    off = _offset_text(dt.datetime.fromtimestamp(ts, tz).utcoffset().total_seconds())
    key = getattr(tz, 'key', None)
    if key and key != 'localtime':
        return '%s, %s' % (key, off)
    return 'UTC' if off == 'UTC+00:00' else off


def tz_short(tz, ts):
    """例: 'JST'（略称がなければ 'UTC+09:00'）"""
    d = dt.datetime.fromtimestamp(ts, tz)
    name = d.tzname() or ''
    if re.fullmatch(r'[A-Z]{2,5}', name):
        return name
    off = _offset_text(d.utcoffset().total_seconds())
    return 'UTC' if off == 'UTC+00:00' else off


def fmt_dt(ts, tz, ms=False):
    d = dt.datetime.fromtimestamp(ts, tz)
    s = d.strftime('%Y/%m/%d %H:%M:%S')
    return s + ('.%03d' % (d.microsecond // 1000)) if ms else s


class DisplayOffset:
    """Excel に表示するタイムゾーンの UTC オフセット（秒）。15 分単位でキャッシュする"""

    def __init__(self, tz):
        self.tz = tz
        self.fixed = tz.utcoffset(None).total_seconds() if isinstance(tz, dt.timezone) else None
        self.cache = {}

    def __call__(self, ts):
        if self.fixed is not None:
            return self.fixed
        key = int(ts // 900)
        off = self.cache.get(key)
        if off is None:
            off = dt.datetime.fromtimestamp(key * 900, self.tz).utcoffset().total_seconds()
            self.cache[key] = off
        return off

    def serial(self, ts):
        """UNIX 時刻を Excel のシリアル値（表示タイムゾーンの日時）に変換する"""
        return 25569.0 + (ts + self(ts)) / 86400.0


class LocalOffset:
    """タイムゾーン表記のないログ日時（ログ側の現地時刻）の UTC オフセット（秒）"""

    def __init__(self, tz):
        self.tz = tz
        self.fixed = tz.utcoffset(None).total_seconds() if isinstance(tz, dt.timezone) else None
        self.cache = {}

    def __call__(self, Y, m, d, H, M):
        if self.fixed is not None:
            return self.fixed
        key = (Y, m, d, H, M)
        off = self.cache.get(key)
        if off is None:
            off = self.tz.utcoffset(dt.datetime(Y, m, d, H, M)).total_seconds()
            if len(self.cache) > 200000:
                self.cache.clear()
            self.cache[key] = off
        return off


# ---- 利用者が指定する日時（--from / --to）
_UT_TZ = r'\s*(?P<tz>Z|UTC|GMT|[+-]\d{2}:?\d{2}|[+-]\d{2})?'
_UT_TIME = r'(?P<H>\d{1,2})(?::(?P<M>\d{2})(?::(?P<S>\d{2})(?:[.,](?P<f>\d{1,6}))?)?)?'
_UT_DT = re.compile(
    r'^(?P<Y>\d{4})(?P<sep>[-/.]?)(?P<m>\d{1,2})(?P=sep)(?P<d>\d{1,2})'
    r'(?:(?:[T\s]+|_)(?:' + _UT_TIME + r'|(?P<cH>\d{2})(?P<cM>\d{2})(?P<cS>\d{2})?))?' + _UT_TZ + r'$', re.I)
_UT_TIME_ONLY = re.compile(r'^(?P<H>\d{1,2}):(?P<M>\d{2})(?::(?P<S>\d{2})(?:[.,](?P<f>\d{1,6}))?)?' + _UT_TZ + r'$',
                           re.I)
_UT_REL = re.compile(r'^(?:now\s*)?-\s*(\d+(?:\.\d+)?)\s*([a-z]+)$', re.I)
_UT_UNITS = {}
for _names, _sec in ((('s', 'sec', 'secs', 'second', 'seconds'), 1),
                     (('m', 'min', 'mins', 'minute', 'minutes'), 60),
                     (('h', 'hr', 'hrs', 'hour', 'hours'), 3600),
                     (('d', 'day', 'days'), 86400),
                     (('w', 'week', 'weeks'), 604800)):
    for _n in _names:
        _UT_UNITS[_n] = _sec


def _user_tz(spec, default):
    if not spec:
        return default
    off = tz_offset(spec)
    if off is None:
        raise ValueError(spec)
    return dt.timezone(dt.timedelta(seconds=off)) if off else UTC


def parse_user_time(text, tz, now):
    """利用者が指定した日時を (開始時刻, 精度) で返す。
    終了日時は「開始時刻 + 精度」を含まない上限として扱う（指定した精度の終わりまでを含む）。"""
    s = text.strip()
    low = s.lower()
    tiny = dt.timedelta(microseconds=1)
    if low == 'now':
        return now, tiny
    if low in ('today', 'yesterday'):
        d = now.astimezone(tz).date() - dt.timedelta(days=1 if low == 'yesterday' else 0)
        return dt.datetime(d.year, d.month, d.day, tzinfo=tz), dt.timedelta(days=1)
    m = _UT_REL.match(s)
    if m:
        unit = _UT_UNITS.get(m.group(2).lower())
        if unit is None:
            raise ValueError(text)
        return now - dt.timedelta(seconds=float(m.group(1)) * unit), tiny
    m = _UT_DT.match(s)
    if m:
        g = m.groupdict()
        Y, mo, d = int(g['Y']), int(g['m']), int(g['d'])
        if g['H'] is not None:
            H, M, S, f = g['H'], g['M'], g['S'], g['f']
        elif g['cH'] is not None:
            H, M, S, f = g['cH'], g['cM'], g['cS'], None
        else:
            H = M = S = f = None
        zone = _user_tz(g['tz'], tz)
    else:
        m = _UT_TIME_ONLY.match(s)
        if not m:
            raise ValueError(text)
        g = m.groupdict()
        zone = _user_tz(g['tz'], tz)
        today = now.astimezone(zone).date()
        Y, mo, d = today.year, today.month, today.day
        H, M, S, f = g['H'], g['M'], g['S'], g['f']
    if H is None:
        step, H, M, S, us = dt.timedelta(days=1), 0, 0, 0, 0
    else:
        if M is None:
            step = dt.timedelta(hours=1)
        elif S is None:
            step = dt.timedelta(minutes=1)
        elif f is None:
            step = dt.timedelta(seconds=1)
        else:
            step = dt.timedelta(microseconds=10 ** (6 - len(f)))
        us = int(f.ljust(6, '0')) if f else 0
        H, M, S = int(H), int(M or 0), int(S or 0)
    return dt.datetime(Y, mo, d, H, M, S, us, tzinfo=zone), step


# =============================================================================
# ログの日時の判定
# =============================================================================
_MONTH = {n: i for i, n in enumerate(
    ('jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'), 1)}
_MON = r'(?P<mon>(?i:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec))'
_FRAC = r'(?:[.,](?P<f>\d{1,9}))?'
_TZ = r'(?:(?P<tz>Z|[+-]\d{2}:?\d{2}|[+-]\d{2}(?![\d:]))|\s(?P<tz2>UTC|GMT|JST|[+-]\d{2}:?\d{2})(?![\w:]))?'


class TsContext:
    """1 ファイル分の日時解釈の文脈（ログ側のタイムゾーンと、年の推定に使う基準日時）"""

    def __init__(self, local_offset, ref_ts):
        self.off = local_offset
        self.ref = ref_ts
        self.ref_year = dt.datetime.fromtimestamp(ref_ts, local_offset.tz).year
        self.days = {}


def to_epoch(ctx, Y, m, d, H, M, S, frac, tz):
    if Y < 1970 or Y > 2199 or H > 23 or M > 59 or S > 60:
        return None
    key = (Y, m, d)
    days = ctx.days.get(key)
    if days is None:
        try:
            days = dt.date(Y, m, d).toordinal() - 719163
        except ValueError:
            return None
        ctx.days[key] = days
    t = days * 86400 + H * 3600 + M * 60 + (59 if S == 60 else S)
    if frac:
        t += int(frac[:6].ljust(6, '0')) / 1e6
    if tz:
        off = tz_offset(tz)
        if off is None:
            return None
    else:
        off = ctx.off(Y, m, d, H, M)
    return t - off


def _c_ymd(m, ctx):
    Y, mo, d, H, M, S, f, tz, tz2 = m.group('Y', 'm', 'd', 'H', 'M', 'S', 'f', 'tz', 'tz2')
    return to_epoch(ctx, int(Y), int(mo), int(d), int(H), int(M), int(S), f, tz or tz2)


def _c_named(m, ctx):
    g = m.groupdict()
    month = _MONTH[g['mon'].lower()] if g.get('mon') else int(g['m'])
    H = int(g['H'])
    ap = g.get('ap')
    if ap:
        if not 1 <= H <= 12:
            return None
        H = H % 12 + (12 if ap.upper() == 'PM' else 0)
    rest = (int(g['d']), H, int(g['M']), int(g['S']), g.get('f'), g.get('tz') or g.get('tz2'))
    if g.get('Y'):
        return to_epoch(ctx, int(g['Y']), month, *rest)
    # 年の記載がない形式（syslog など）は、最終更新日時から年を推定する
    t = to_epoch(ctx, ctx.ref_year, month, *rest)
    if t is None or t > ctx.ref + 2 * 86400:
        t = to_epoch(ctx, ctx.ref_year - 1, month, *rest)
    return t


def _c_epoch(m, ctx):
    s = m.group('num')
    digits = len(s.split('.', 1)[0])
    scale = {10: 1.0, 13: 1e3, 16: 1e6, 19: 1e9}.get(digits)
    if scale is None:
        return None
    v = float(s) / scale
    return v if 946684800 <= v <= 7258118400 else None


TS_PATTERNS = [(name, label, re.compile(rx), conv) for name, label, rx, conv in (
    ('iso', 'ISO 8601',
     r'(?<!\d)(?P<Y>\d{4})-(?P<m>\d{2})-(?P<d>\d{2})[T \t](?P<H>\d{2}):(?P<M>\d{2}):(?P<S>\d{2})' + _FRAC + _TZ,
     _c_ymd),
    ('slash', 'YYYY/MM/DD hh:mm:ss',
     r'(?<!\d)(?P<Y>\d{4})/(?P<m>\d{1,2})/(?P<d>\d{1,2})[T \t]+(?P<H>\d{1,2}):(?P<M>\d{2}):(?P<S>\d{2})' + _FRAC + _TZ,
     _c_ymd),
    ('clf', 'Apache / Nginx',
     r'(?<!\d)(?P<d>\d{1,2})/' + _MON + r'/(?P<Y>\d{4}):(?P<H>\d{2}):(?P<M>\d{2}):(?P<S>\d{2})' + _FRAC +
     r'(?:\s(?P<tz>[+-]\d{4}))?',
     _c_named),
    ('tomcat', 'DD-Mon-YYYY hh:mm:ss',
     r'(?<!\d)(?P<d>\d{1,2})-' + _MON + r'-(?P<Y>\d{4})[ T:](?P<H>\d{2}):(?P<M>\d{2}):(?P<S>\d{2})' + _FRAC + _TZ,
     _c_named),
    ('java', 'Mon DD, YYYY h:mm:ss AM',
     r'(?<![A-Za-z])' + _MON + r'\s+(?P<d>\d{1,2}),\s+(?P<Y>\d{4})\s+(?P<H>\d{1,2}):(?P<M>\d{2}):(?P<S>\d{2})'
     r'\s*(?P<ap>(?i:AM|PM))',
     _c_named),
    ('syslog', 'syslog',
     r'(?<![A-Za-z])' + _MON + r'\s+(?P<d>\d{1,2})\s+(?P<H>\d{2}):(?P<M>\d{2}):(?P<S>\d{2})' + _FRAC +
     r'(?:\s+(?P<Y>\d{4})(?![\d:]))?',
     _c_named),
    ('jp', 'YYYY年MM月DD日 hh時mm分ss秒',
     r'(?P<Y>\d{4})年\s*(?P<m>\d{1,2})月\s*(?P<d>\d{1,2})日[^\d]{0,8}?(?P<H>\d{1,2})[時:](?P<M>\d{1,2})[分:]'
     r'(?P<S>\d{1,2})秒?',
     _c_named),
    ('isobasic', 'YYYYMMDDThhmmss',
     r'(?<!\d)(?P<Y>\d{4})(?P<m>\d{2})(?P<d>\d{2})T(?P<H>\d{2})(?P<M>\d{2})(?P<S>\d{2})' + _FRAC + _TZ,
     _c_ymd),
    ('json-epoch', 'JSON（エポック時刻）',
     r'"(?i:@?timestamp|time|ts|datetime|date|eventtime|event_time|logged_at|created_at|timemillis|epoch)"'
     r'\s*:\s*"?(?P<num>\d{10}(?:\.\d{1,9})?|\d{13}(?:\.\d{1,6})?|\d{16}|\d{19})(?![\d.])',
     _c_epoch),
    ('epoch', 'エポック時刻（行頭）',
     r'^(?P<num>\d{10}(?:\.\d{1,9})?|\d{13})(?=[\s,;|]|$)',
     _c_epoch),
)]


class LineTimeParser:
    """判定した形式で、各行の先頭付近から日時を読み取る"""
    __slots__ = ('search', 'conv', 'ctx', 'window')

    def __init__(self, rx, conv, ctx, window):
        self.search, self.conv, self.ctx, self.window = rx.search, conv, ctx, window

    def __call__(self, line):
        m = self.search(line, 0, self.window)
        return None if m is None else self.conv(m, self.ctx)


def detect_time_parser(lines, ctx):
    """先頭の行を各形式で試し、最も多くの行から日時を読み取れた形式を選ぶ"""
    best = None
    for name, label, rx, conv in TS_PATTERNS:
        hits = max_off = 0
        search = rx.search
        for line in lines:
            m = search(line, 0, 1024)
            if m is not None and conv(m, ctx) is not None:
                hits += 1
                if m.start() > max_off:
                    max_off = m.start()
        if hits and (best is None or hits > best[0]):
            best = (hits, label, rx, conv, max_off)
    if best is None:
        return None, None
    hits, label, rx, conv, max_off = best
    return LineTimeParser(rx, conv, ctx, min(1024, max(96, max_off + 96))), label


class CustomTimeSpec:
    """--ts-regex / --ts-format による日時の指定"""

    def __init__(self, regex, fmt):
        try:
            self.rx = re.compile(regex)
        except re.error as e:
            raise UsageError('--ts-regex の正規表現が正しくありません: %s（%s）' % (regex, e))
        self.group = 'ts' if 'ts' in self.rx.groupindex else (1 if self.rx.groups else 0)
        f = (fmt or 'auto').strip()
        low = f.lower()
        self.fmt = f
        self.scale = {'epoch': 1.0, 'epoch_s': 1.0, 'epoch_ms': 1e3, 'epoch_us': 1e6, 'epoch_ns': 1e9}.get(low)
        if self.scale is not None:
            self.kind = 'epoch'
        elif low == 'auto':
            self.kind = 'auto'
        else:
            self.kind = 'strptime'
            if '%' not in f:
                raise UsageError('--ts-format に書式（%%Y-%%m-%%d %%H:%%M:%%S など）を指定してください: %s' % f)
            try:
                dt.datetime.now().strftime(f)
            except ValueError as e:
                raise UsageError('--ts-format の書式が正しくありません: %s（%s）' % (f, e))
            self.no_year = not re.search(r'%[YyGC]', f)
        self.label = 'カスタム（%s）' % (f if self.kind != 'auto' else '--ts-regex')

    def bind(self, ctx):
        spec = self
        rx_search = spec.rx.search

        def parse(line):
            m = rx_search(line)
            if m is None:
                return None
            s = m.group(spec.group)
            if not s:
                return None
            s = s.strip()
            if spec.kind == 'epoch':
                try:
                    return float(s) / spec.scale
                except ValueError:
                    return None
            if spec.kind == 'auto':
                for _name, _label, rx, conv in TS_PATTERNS:
                    mm = rx.search(s)
                    if mm is not None:
                        v = conv(mm, ctx)
                        if v is not None:
                            return v
                return None
            try:
                d = dt.datetime.strptime(s, spec.fmt)
            except ValueError:
                return None
            if d.tzinfo is not None:
                return d.timestamp()
            frac = '%06d' % d.microsecond if d.microsecond else None
            if not spec.no_year:
                return to_epoch(ctx, d.year, d.month, d.day, d.hour, d.minute, d.second, frac, None)
            t = to_epoch(ctx, ctx.ref_year, d.month, d.day, d.hour, d.minute, d.second, frac, None)
            if t is None or t > ctx.ref + 2 * 86400:
                t = to_epoch(ctx, ctx.ref_year - 1, d.month, d.day, d.hour, d.minute, d.second, frac, None)
            return t
        return parse


# =============================================================================
# 強調表示とテキストの整形
# =============================================================================
class Highlighter:
    def __init__(self, keywords, ignore_case, regex):
        self.keywords = keywords
        flags = re.IGNORECASE if ignore_case else 0
        order = list(range(len(keywords)))
        if not regex:
            # 同じ位置で複数の文言が一致するときは長い方を優先する
            order.sort(key=lambda i: -len(keywords[i]))
        parts = []
        for i in order:
            pat = keywords[i] if regex else re.escape(keywords[i])
            try:
                re.compile(pat, flags)
            except re.error as e:
                raise UsageError('強調文言の正規表現が正しくありません: %s（%s）' % (keywords[i], e))
            parts.append('(?P<_kw%d>%s)' % (i, pat))
        self.rx = re.compile('|'.join(parts), flags)

    def find(self, text):
        out = []
        for m in self.rx.finditer(text):
            a, b = m.span()
            if a != b:
                out.append((a, b, int(m.lastgroup[3:])))
        return out


_ANSI_RE = re.compile(r'\x1b(?:\[[0-?]*[ -/]*[@-~]|\][^\x07\x1b]*(?:\x07|\x1b\\)?|[@-Z\\-_])')
_CTRL_RE = re.compile('[\x00-\x08\x0b-\x1f\x7f-\x9f\ufffe\uffff\ud800-\udfff]')


def clean_text(s):
    """エスケープシーケンス（色付け）と制御文字を取り除き、タブを空白にする"""
    if '\x1b' in s:
        s = _ANSI_RE.sub('', s)
    if _CTRL_RE.search(s):
        s = _CTRL_RE.sub('', s)
    if '\t' in s:
        s = s.expandtabs(4)
    return s


def truncate_cell(s):
    """Excel の 1 セルの上限（32,767 文字）に収まるように切り詰める"""
    if len(s.encode('utf-16-le')) // 2 <= CELL_LIMIT:
        return s
    keep = CELL_LIMIT - 40
    t = s[:keep]
    over = len(t.encode('utf-16-le')) // 2 - keep
    while over > 0:
        t = t[:-over]
        over = len(t.encode('utf-16-le')) // 2 - keep
    return t + '…［以下 %s 文字省略］' % fmt_n(len(s) - len(t))


def text_width(s):
    """Excel の列幅の目安（半角 1・全角 2）"""
    if s.isascii():
        return len(s)
    return (len(s.encode('utf-8')) + len(s)) // 2


# =============================================================================
# 圧縮形式・文字コードの判定と読み込み
# =============================================================================
_READ_ERRORS = (EOFError, OSError, ValueError, lzma.LZMAError, zlib.error)


def sniff_compression(path):
    with open(path, 'rb') as f:
        head = f.read(6)
    if head[:2] == b'\x1f\x8b':
        return 'gzip'
    if head[:3] == b'BZh' and head[3:4].isdigit():
        return 'bzip2'
    if head[:6] == b'\xfd7zXZ\x00':
        return 'xz'
    if head[:4] == b'PK\x03\x04':
        return 'zip'
    return None


def open_binary(path, comp):
    if comp == 'gzip':
        return gzip.open(path, 'rb')
    if comp == 'bzip2':
        return bz2.open(path, 'rb')
    if comp == 'xz':
        return lzma.open(path, 'rb')
    return open(path, 'rb')


_BOMS = ((codecs.BOM_UTF8, 'utf-8-sig'), (codecs.BOM_UTF16_LE, 'utf-16'), (codecs.BOM_UTF16_BE, 'utf-16'))
_ENC_ALIASES = {'sjis': 'cp932', 'shift_jis': 'cp932', 'shift-jis': 'cp932', 'shiftjis': 'cp932', 'ms932': 'cp932',
                'windows-31j': 'cp932', 'eucjp': 'euc_jp', 'euc-jp': 'euc_jp', 'utf8': 'utf-8-sig',
                'utf-8': 'utf-8-sig', 'utf_8': 'utf-8-sig'}
_ENC_LABELS = {'utf-8': 'UTF-8', 'utf-8-sig': 'UTF-8', 'cp932': 'Shift_JIS', 'shift_jis': 'Shift_JIS',
               'euc_jp': 'EUC-JP', 'utf-16': 'UTF-16', 'utf-16-le': 'UTF-16', 'utf-16-be': 'UTF-16'}


def normalize_encoding(name):
    if not name:
        return None
    n = _ENC_ALIASES.get(name.strip().lower(), name.strip())
    try:
        codecs.lookup(n)
    except LookupError:
        raise UsageError('--encoding に指定した文字コードが不明です: %s' % name)
    return n


def encoding_label(enc):
    try:
        return _ENC_LABELS.get(codecs.lookup(enc).name, enc.upper())
    except LookupError:
        return enc


def detect_encoding(path, comp):
    """BOM → UTF-8 → EUC-JP → Shift_JIS の順に判定する。テキストでなければ None を返す"""
    with open_binary(path, comp) as f:
        chunk = f.read(1 << 16)
        for bom, name in _BOMS:
            if chunk.startswith(bom):
                return name
        if b'\x00' in chunk[:8192]:
            return None
        cands = [(n, codecs.getincrementaldecoder(n)()) for n in ('utf-8', 'euc_jp', 'cp932')]
        total = 0
        while True:
            if not chunk:
                alive = []
                for n, dec in cands:
                    try:
                        dec.decode(b'', True)
                        alive.append((n, dec))
                    except UnicodeDecodeError:
                        pass
                cands = alive
                break
            alive = []
            for n, dec in cands:
                try:
                    dec.decode(chunk)
                    alive.append((n, dec))
                except UnicodeDecodeError:
                    pass
            cands = alive
            total += len(chunk)
            if len(cands) <= 1 or total >= DETECT_LIMIT:
                break
            try:
                chunk = f.read(1 << 20)
            except _READ_ERRORS:
                break
    return cands[0][0] if cands else 'utf-8'


class ReadState:
    def __init__(self):
        self.error = None


def iter_lines(path, comp, enc, state):
    """1 行ずつ返す（改行は除く）。圧縮データが途中で壊れていたら state.error に記録して終わる"""
    text = io.TextIOWrapper(open_binary(path, comp), encoding=enc, errors='replace', newline=None)
    try:
        for line in text:
            yield line[:-1] if line.endswith('\n') else line
    except _READ_ERRORS as e:
        state.error = str(e) or e.__class__.__name__
    finally:
        text.close()


# =============================================================================
# Excel（Office Open XML）の書き出し
# =============================================================================
NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

_XESC_RE = re.compile(r'_(?=x[0-9A-Fa-f]{4}_)')


def _x(s):
    if '&' in s:
        s = s.replace('&', '&amp;')
    if '<' in s:
        s = s.replace('<', '&lt;')
    if '>' in s:
        s = s.replace('>', '&gt;')
    return s


def _xa(s):
    return _x(s).replace('"', '&quot;')


def _xs(s):
    """セルの文字列用（_x0041_ のような文字列が Excel で変換されないようにする）"""
    if '_x' in s:
        s = _XESC_RE.sub('_x005F_', s)
    return _x(s)


def _num(v):
    return ('%d' % v) if float(v).is_integer() else ('%g' % v)


def col_letter(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def quote_sheet(name):
    return "'%s'" % name.replace("'", "''")


def rpr(size=10, color=C_TEXT, bold=False):
    return '<rPr>%s<sz val="%s"/><color rgb="FF%s"/><rFont val="%s"/><family val="3"/><charset val="128"/></rPr>' % (
        '<b/>' if bold else '', _num(size), color, FONT_NAME)


def rich(runs):
    return ''.join('<r>%s<t xml:space="preserve">%s</t></r>' % (rp, _xs(t)) for t, rp in runs if t)


def c_str(ref, s, text):
    if not text:
        return '<c r="%s" s="%d"/>' % (ref, s)
    return '<c r="%s" s="%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (ref, s, _xs(text))


def c_rich(ref, s, runs):
    body = rich(runs)
    if not body:
        return '<c r="%s" s="%d"/>' % (ref, s)
    return '<c r="%s" s="%d" t="inlineStr"><is>%s</is></c>' % (ref, s, body)


def c_num(ref, s, v):
    return '<c r="%s" s="%d"><v>%s</v></c>' % (ref, s, ('%d' % v) if isinstance(v, int) else repr(float(v)))


def c_blank(ref, s):
    return '<c r="%s" s="%d"/>' % (ref, s)


def row_xml(r, cells, ht=None):
    attr = (' ht="%s" customHeight="1"' % _num(ht)) if ht else ''
    return '<row r="%d"%s>%s</row>' % (r, attr, ''.join(cells))


class StyleBook:
    """styles.xml の内容（フォント・塗りつぶし・罫線・表示形式・セルの書式）を重複なく管理する"""
    _BUILTIN_NUMFMT = {'General': 0, '0': 1, '#,##0': 3}

    def __init__(self):
        self._fonts, self._fills, self._borders, self._xfs = [], [], [], []
        self._index = {}
        self._numfmts = collections.OrderedDict()
        self.font()     # 0: 既定のフォント（標準スタイル）。空のセルも Meiryo UI になる
        self._fills += ['<fill><patternFill patternType="none"/></fill>',
                        '<fill><patternFill patternType="gray125"/></fill>']
        self.border()   # 0: 罫線なし
        self.xf()       # 0: 既定のセルの書式

    def _intern(self, table, kind, xml):
        idx = self._index.get((kind, xml))
        if idx is None:
            idx = self._index[(kind, xml)] = len(table)
            table.append(xml)
        return idx

    def font(self, size=10, color=C_TEXT, bold=False, underline=False):
        return self._intern(self._fonts, 'font', '<font>%s%s<sz val="%s"/><color rgb="FF%s"/><name val="%s"/>'
                            '<family val="3"/><charset val="128"/></font>' % (
                                '<b/>' if bold else '', '<u/>' if underline else '', _num(size), color, FONT_NAME))

    def fill(self, rgb=None):
        if rgb is None:
            return 0
        return self._intern(self._fills, 'fill', '<fill><patternFill patternType="solid"><fgColor rgb="FF%s"/>'
                            '<bgColor indexed="64"/></patternFill></fill>' % rgb)

    def border(self, left=None, right=None, top=None, bottom=None):
        def side(tag, spec):
            if not spec:
                return '<%s/>' % tag
            return '<%s style="%s"><color rgb="FF%s"/></%s>' % (tag, spec[0], spec[1], tag)
        return self._intern(self._borders, 'border', '<border>%s%s%s%s<diagonal/></border>' % (
            side('left', left), side('right', right), side('top', top), side('bottom', bottom)))

    def numfmt(self, code):
        if code in self._BUILTIN_NUMFMT:
            return self._BUILTIN_NUMFMT[code]
        if code not in self._numfmts:
            self._numfmts[code] = 164 + len(self._numfmts)
        return self._numfmts[code]

    def xf(self, font=0, fill=0, border=0, numfmt='General', h=None, v=None, wrap=False, indent=0):
        nf = self.numfmt(numfmt)
        al = []
        if h:
            al.append('horizontal="%s"' % h)
        if v:
            al.append('vertical="%s"' % v)
        if wrap:
            al.append('wrapText="1"')
        if indent:
            al.append('indent="%d"' % indent)
        attrs = '<xf numFmtId="%d" fontId="%d" fillId="%d" borderId="%d" xfId="0"' % (nf, font, fill, border)
        attrs += ''.join(flag for cond, flag in ((nf, ' applyNumberFormat="1"'), (font, ' applyFont="1"'),
                                                 (fill, ' applyFill="1"'), (border, ' applyBorder="1"'),
                                                 (al, ' applyAlignment="1"')) if cond)
        xml = attrs + ('><alignment %s/></xf>' % ' '.join(al) if al else '/>')
        return self._intern(self._xfs, 'xf', xml)

    def xml(self):
        out = [XML_HEAD, '<styleSheet xmlns="%s">' % NS_MAIN]
        if self._numfmts:
            out.append('<numFmts count="%d">%s</numFmts>' % (len(self._numfmts), ''.join(
                '<numFmt numFmtId="%d" formatCode="%s"/>' % (i, _xa(code)) for code, i in self._numfmts.items())))
        out.append('<fonts count="%d">%s</fonts>' % (len(self._fonts), ''.join(self._fonts)))
        out.append('<fills count="%d">%s</fills>' % (len(self._fills), ''.join(self._fills)))
        out.append('<borders count="%d">%s</borders>' % (len(self._borders), ''.join(self._borders)))
        out.append('<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>')
        out.append('<cellXfs count="%d">%s</cellXfs>' % (len(self._xfs), ''.join(self._xfs)))
        out.append('<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>')
        out.append('<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium9" '
                   'defaultPivotStyle="PivotStyleLight16"/></styleSheet>')
        return ''.join(out)


class DataStyles:
    """ログのシートで使う書式"""

    def __init__(self, sb, wrap, n_kw):
        v = 'top' if wrap else 'center'
        grid = ('thin', C_GRID)
        bd = sb.border(left=grid, right=grid)
        f_ln, f_ts, f_cont, f_txt = sb.font(color=C_MUTED), sb.font(), sb.font(color=C_FAINT), sb.font()
        fills = [sb.fill(None), sb.fill(C_BAND)] + [sb.fill(KW_PALETTE[k % len(KW_PALETTE)][1]) for k in range(n_kw)]
        # rows[0]: 通常 / rows[1]: 縞模様 / rows[2 + k]: k 番目の強調文言を含む行
        #   各要素は（行番号, 日時, 継続行の日時, ログ）の書式
        self.rows = [(
            sb.xf(font=f_ln, fill=fi, border=bd, h='right', v=v, indent=1),
            sb.xf(font=f_ts, fill=fi, border=bd, numfmt=FMT_TS, h='center', v=v),
            sb.xf(font=f_cont, fill=fi, border=bd, numfmt=FMT_TS, h='center', v=v),
            sb.xf(font=f_txt, fill=fi, border=bd, h='left', v=v, wrap=wrap, indent=1),
        ) for fi in fills]
        hb = sb.border(left=('thin', C_NAVY_LINE), right=('thin', C_NAVY_LINE), top=('thin', C_NAVY),
                       bottom=('thin', C_NAVY))
        hf, hfill = sb.font(bold=True, color=C_WHITE), sb.fill(C_NAVY)
        self.head_c = sb.xf(font=hf, fill=hfill, border=hb, h='center', v='center')
        self.head_l = sb.xf(font=hf, fill=hfill, border=hb, h='left', v='center', indent=1)
        rule = sb.border(bottom=('medium', C_NAVY))
        self.title = sb.xf(font=sb.font(size=14, bold=True, color=C_NAVY), border=rule, v='center')
        self.title_rule = sb.xf(border=rule)
        self.meta = sb.xf(font=sb.font(size=9, color=C_SUB), v='center')


# 情報行（リッチテキスト）の文字書式
RP_LABEL = rpr(9, C_SUB)
RP_VALUE = rpr(9, C_TEXT)
RP_STRONG = rpr(9, C_TEXT, bold=True)
RP_SEP = rpr(9, C_FAINT)
RP_NOTE = rpr(9, STATUS_COLOR['warn'], bold=True)
RP_MUTED = rpr(9, C_MUTED)
SEP = ('　｜　', RP_SEP)


def sheet_head(dim, cols, freeze=None, tab_color=None, selected=False, row_ht=18.0, custom_ht=True,
               fit_width=False):
    parts = [XML_HEAD, '<worksheet xmlns="%s" xmlns:r="%s">' % (NS_MAIN, NS_REL)]
    pr = ('<tabColor rgb="FF%s"/>' % tab_color if tab_color else '') + \
         ('<pageSetUpPr fitToPage="1"/>' if fit_width else '')
    if pr:
        parts.append('<sheetPr>%s</sheetPr>' % pr)
    parts.append('<dimension ref="%s"/>' % dim)
    sv = '<sheetView showGridLines="0"%s zoomScale="100" zoomScaleNormal="100" workbookViewId="0">' % (
        ' tabSelected="1"' if selected else '')
    if freeze:
        row, col = freeze
        tl = '%s%d' % (col_letter(col), row)
        if col > 1:
            sv += ('<pane xSplit="%d" ySplit="%d" topLeftCell="%s" activePane="bottomRight" state="frozen"/>'
                   '<selection pane="topRight"/><selection pane="bottomLeft"/>'
                   '<selection pane="bottomRight" activeCell="%s" sqref="%s"/>') % (col - 1, row - 1, tl, tl, tl)
        else:
            sv += ('<pane ySplit="%d" topLeftCell="%s" activePane="bottomLeft" state="frozen"/>'
                   '<selection pane="bottomLeft" activeCell="%s" sqref="%s"/>') % (row - 1, tl, tl, tl)
    else:
        sv += '<selection activeCell="A1" sqref="A1"/>'
    parts.append('<sheetViews>%s</sheetView></sheetViews>' % sv)
    parts.append('<sheetFormatPr defaultRowHeight="%s"%s/>' % (_num(row_ht), ' customHeight="1"' if custom_ht else ''))
    parts.append('<cols>%s</cols>' % ''.join('<col min="%d" max="%d" width="%s" customWidth="1"/>' % (
        i, i, _num(round(w, 2))) for i, w in enumerate(cols, 1)))
    parts.append('<sheetData>')
    return ''.join(parts)


def sheet_tail(autofilter=None, hyperlinks=(), fit_width=False):
    parts = ['</sheetData>']
    if autofilter:
        parts.append('<autoFilter ref="%s"/>' % autofilter)
    if hyperlinks:
        parts.append('<hyperlinks>%s</hyperlinks>' % ''.join(
            '<hyperlink ref="%s" location="%s" display="%s"/>' % (ref, _xa(loc), _xa(disp))
            for ref, loc, disp in hyperlinks))
    parts.append('<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.6" header="0.3" footer="0.3"/>')
    parts.append('<pageSetup paperSize="9" orientation="landscape"%s/>' % (
        ' fitToWidth="1" fitToHeight="0"' if fit_width else ''))
    parts.append('<headerFooter><oddFooter>&amp;L&amp;8&amp;A&amp;R&amp;8&amp;P / &amp;N</oddFooter></headerFooter>')
    parts.append('</worksheet>')
    return ''.join(parts)


class DataSheet:
    """ログのシート 1 枚分。行は一時ファイルに書き出しておき、最後に Excel ファイルへまとめる"""

    def __init__(self, name, part, order, body_path, n_kw):
        self.name, self.part, self.order, self.body_path = name, part, order, body_path
        self.fh = open(body_path, 'w', encoding='utf-8', newline='\n', buffering=1 << 20)
        self.row = HEADER_ROWS
        self.data_rows = 0
        self.hist = [0] * (WIDTH_MAX + 1)
        self.kw_hits = [0] * n_kw
        self.max_lineno = 0
        self.ts_min = self.ts_max = None

    def log_col_width(self, wrap):
        """ログ列の幅（行の 99% が収まる幅。長すぎる行は折り返し、または右にはみ出して表示）"""
        need, cum, p = self.data_rows * 0.99, 0, WIDTH_MAX
        for w, c in enumerate(self.hist):
            cum += c
            if cum >= need:
                p = w
                break
        return float(min(120 if wrap else 160, max(40, p + 3)))


class XlsxBook:
    def __init__(self, path, workdir, sb, wrap):
        self.path = path
        self.tmp_path = '%s.tmp%d' % (path, os.getpid())
        self.workdir = workdir
        self.sb = sb
        self.wrap = wrap
        self.zf = zipfile.ZipFile(self.tmp_path, 'w', zipfile.ZIP_DEFLATED)
        self.sheets = []          # データシート（ブック内の並び順）
        self.defined = []         # (名前, シート位置, 参照範囲, 非表示)
        self.used = {SUMMARY_SHEET.casefold(), 'history'}
        self.next_part = 2        # sheet1.xml はサマリー
        self.ds = None

    # ---- シート名（31 文字以内・使えない文字なし・重複なし）
    def unique_name(self, base, suffix=''):
        n = 1
        while True:
            tail = suffix + ('' if n == 1 else ' (%d)' % n)
            limit = 31 - len(tail)
            name = base
            if len(name) > limit:
                head = (limit - 1) * 2 // 5          # 末尾（日付など）を多めに残す
                name = name[:head] + '…' + name[len(name) - (limit - 1 - head):]
            name = name.strip().strip("'") + tail
            if name.casefold() not in self.used:
                self.used.add(name.casefold())
                return name
            n += 1

    def release_name(self, name):
        self.used.discard(name.casefold())

    def new_data_sheet(self, name, n_kw):
        part = self.next_part
        self.next_part += 1
        sh = DataSheet(name, part, len(self.sheets) + 1, os.path.join(self.workdir, 'sheet%d.part' % part), n_kw)
        self.sheets.append(sh)
        return sh

    def discard_sheet(self, sh):
        try:
            sh.fh.close()
            os.remove(sh.body_path)
        except OSError:
            pass
        self.sheets.remove(sh)
        self.release_name(sh.name)
        for i, s in enumerate(self.sheets, 1):
            s.order = i

    def _write(self, arcname, head, body_path=None, tail=''):
        hb, tb = head.encode('utf-8'), tail.encode('utf-8')
        zi = zipfile.ZipInfo(arcname, date_time=time.localtime()[:6])
        zi.compress_type = zipfile.ZIP_DEFLATED
        # 事前にサイズを伝えておくと、必要なときだけ ZIP64 拡張が使われる
        zi.file_size = len(hb) + len(tb) + (os.path.getsize(body_path) if body_path else 0)
        with self.zf.open(zi, 'w') as w:
            w.write(hb)
            if body_path:
                with open(body_path, 'rb') as f:
                    shutil.copyfileobj(f, w, 1 << 20)
            w.write(tb)

    def finalize_data_sheet(self, sh, header_rows, tab_color):
        last = sh.row
        cols = [max(9.0, len(str(sh.max_lineno)) + 4.5), 25.0, sh.log_col_width(self.wrap)]
        head = sheet_head('A1:C%d' % last, cols, freeze=(HEADER_ROWS + 1, 3), tab_color=tab_color,
                          row_ht=14.4 if self.wrap else ROW_HT, custom_ht=not self.wrap) + header_rows
        tail = sheet_tail(autofilter='A%d:C%d' % (HEADER_ROWS, last))
        self._write('xl/worksheets/sheet%d.xml' % sh.part, head, sh.body_path, tail)
        os.remove(sh.body_path)
        q = quote_sheet(sh.name)
        self.defined.append(('_xlnm._FilterDatabase', sh.order, '%s!$A$%d:$C$%d' % (q, HEADER_ROWS, last), True))
        self.defined.append(('_xlnm.Print_Titles', sh.order, '%s!$%d:$%d' % (q, HEADER_ROWS, HEADER_ROWS), False))

    def write_summary(self, xml, filter_ref, title_row):
        self._write('xl/worksheets/sheet1.xml', xml)
        q = quote_sheet(SUMMARY_SHEET)
        if filter_ref:
            self.defined.append(('_xlnm._FilterDatabase', 0, '%s!%s' % (q, filter_ref), True))
        self.defined.append(('_xlnm.Print_Titles', 0, '%s!$%d:$%d' % (q, title_row, title_row), False))

    def close(self, title, created):
        entries = [(SUMMARY_SHEET, 1)] + [(sh.name, sh.part) for sh in self.sheets]
        wb = [XML_HEAD, '<workbook xmlns="%s" xmlns:r="%s"><workbookPr/>' % (NS_MAIN, NS_REL),
              '<bookViews><workbookView xWindow="240" yWindow="120" windowWidth="28800" windowHeight="16000" '
              'tabRatio="780" firstSheet="0" activeTab="0"/></bookViews><sheets>']
        for i, (name, _part) in enumerate(entries, 1):
            wb.append('<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (_xa(name), i, i))
        wb.append('</sheets>')
        if self.defined:
            wb.append('<definedNames>%s</definedNames>' % ''.join(
                '<definedName name="%s" localSheetId="%d"%s>%s</definedName>' % (
                    n, lid, ' hidden="1"' if hidden else '', _x(ref))
                for n, lid, ref, hidden in sorted(self.defined, key=lambda d: (d[0], d[1]))))
        wb.append('</workbook>')
        self._write('xl/workbook.xml', ''.join(wb))

        rels = [XML_HEAD, '<Relationships xmlns="%s">' % NS_PKG_REL]
        for i, (_name, part) in enumerate(entries, 1):
            rels.append('<Relationship Id="rId%d" Type="%s/worksheet" Target="worksheets/sheet%d.xml"/>' % (
                i, NS_REL, part))
        rels.append('<Relationship Id="rId%d" Type="%s/styles" Target="styles.xml"/></Relationships>' % (
            len(entries) + 1, NS_REL))
        self._write('xl/_rels/workbook.xml.rels', ''.join(rels))
        self._write('xl/styles.xml', self.sb.xml())

        stamp = created.astimezone(UTC).strftime('%Y-%m-%dT%H:%M:%SZ')
        self._write('docProps/core.xml', XML_HEAD + (
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            '<dc:title>%s</dc:title><dc:creator>s3log2xlsx</dc:creator>'
            '<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created>'
            '<dcterms:modified xsi:type="dcterms:W3CDTF">%s</dcterms:modified></cp:coreProperties>') % (
                _x(title), stamp, stamp))
        self._write('docProps/app.xml', XML_HEAD + (
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
            '<Application>s3log2xlsx</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>'
            '</Properties>'))
        self._write('_rels/.rels', XML_HEAD + (
            '<Relationships xmlns="%s">'
            '<Relationship Id="rId1" Type="%s/officeDocument" Target="xl/workbook.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/'
            'core-properties" Target="docProps/core.xml"/>'
            '<Relationship Id="rId3" Type="%s/extended-properties" Target="docProps/app.xml"/>'
            '</Relationships>') % (NS_PKG_REL, NS_REL, NS_REL))
        ct = [XML_HEAD, '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
              '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
              '<Default Extension="xml" ContentType="application/xml"/>',
              '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.'
              'spreadsheetml.sheet.main+xml"/>']
        for _name, part in entries:
            ct.append('<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.'
                      'openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % part)
        ct.append('<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.'
                  'spreadsheetml.styles+xml"/>')
        ct.append('<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.'
                  'core-properties+xml"/>')
        ct.append('<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-'
                  'officedocument.extended-properties+xml"/></Types>')
        self._write('[Content_Types].xml', ''.join(ct))
        self.zf.close()
        os.replace(self.tmp_path, self.path)

    def abort(self):
        try:
            self.zf.close()
        except Exception:
            pass
        for p in [self.tmp_path] + [sh.body_path for sh in self.sheets]:
            try:
                os.remove(p)
            except OSError:
                pass


# =============================================================================
# 1 ファイル分の抽出結果をシートに書き出す
# =============================================================================
_SHEET_BAD = re.compile(r'[\x00-\x1f\x7f:\\/?*\[\]]')


def sheet_base(key):
    name = key.rstrip('/').rsplit('/', 1)[-1] or 'log'
    low = name.lower()
    for ext in ('.gz', '.gzip', '.bz2', '.xz'):
        if low.endswith(ext) and len(name) > len(ext):
            name = name[:-len(ext)]
            break
    return _SHEET_BAD.sub('_', name).strip().strip("'") or 'log'


class S3Object:
    __slots__ = ('key', 'size', 'mtime')

    def __init__(self, key, size, mtime):
        self.key, self.size, self.mtime = key, size, mtime


class FileResult:
    def __init__(self, obj):
        self.obj = obj
        self.status = 'ok'
        self.message = ''
        self.encoding = None
        self.compression = None
        self.ts_label = None
        self.total_lines = None
        self.matched = 0
        self.kw_hits = None
        self.sheet_names = []
        self.untimed_all = False


class FileSheets:
    """1 つの S3 オブジェクトに対応するシート（行数が上限を超えたら続きのシートを作る）"""

    def __init__(self, book, obj, cfg, res):
        self.book, self.obj, self.cfg, self.res = book, obj, cfg, res
        self.hl = cfg.highlighter
        self.n_kw = len(cfg.keywords)
        self.rows = book.ds.rows
        self.disp = cfg.display_offset
        self.base = sheet_base(obj.key)
        self.sheets = []
        self.sheet = None
        self.rp_base = rpr(10, C_TEXT)
        self.rp_kw = [rpr(10, KW_PALETTE[k % len(KW_PALETTE)][0], bold=True) for k in range(self.n_kw)]
        # 行の高さ。Excel は高さの指定がない行をフォントに合わせて詰めて表示するため明示する。
        # 折り返し表示のときは、Excel が内容に合わせて高さを自動調整するよう指定しない。
        self.row_open = '<row r="%d">' if cfg.wrap else '<row r="%%d" ht="%s" customHeight="1">' % _num(ROW_HT)

    def _new_sheet(self):
        n = len(self.sheets) + 1
        name = self.book.unique_name(self.base, '' if n == 1 else ' #%d' % n)
        self.sheet = self.book.new_data_sheet(name, self.n_kw)
        self.sheets.append(self.sheet)
        return self.sheet

    def add(self, lineno, ts, cont, line, band):
        sh = self.sheet
        if sh is None or sh.data_rows >= self.cfg.max_rows:
            sh = self._new_sheet()
        r = sh.row = sh.row + 1
        sh.data_rows += 1
        sh.max_lineno = lineno
        text = clean_text(line)
        if len(text) > 16383:
            text = truncate_cell(text)
        hits = self.hl.find(text) if self.hl is not None else None
        if hits:
            ks = {k for _a, _b, k in hits}
            for k in ks:
                sh.kw_hits[k] += 1
            st = self.rows[2 + min(ks)]
        else:
            st = self.rows[band]
        w = len(text) if text.isascii() else (len(text.encode('utf-8')) + len(text)) // 2
        sh.hist[w if w < WIDTH_MAX else WIDTH_MAX] += 1
        if ts is None:
            c_ts = '<c r="B%d" s="%d"/>' % (r, st[1])
        else:
            if not cont:
                if sh.ts_min is None or ts < sh.ts_min:
                    sh.ts_min = ts
                if sh.ts_max is None or ts > sh.ts_max:
                    sh.ts_max = ts
            c_ts = '<c r="B%d" s="%d"><v>%r</v></c>' % (r, st[2] if cont else st[1], self.disp.serial(ts))
        if not text:
            c_tx = '<c r="C%d" s="%d"/>' % (r, st[3])
        elif hits:
            c_tx = '<c r="C%d" s="%d" t="inlineStr"><is>%s</is></c>' % (r, st[3], self._runs(text, hits))
        else:
            c_tx = '<c r="C%d" s="%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (
                r, st[3], _xs(text))
        sh.fh.write(self.row_open % r + '<c r="A%d" s="%d"><v>%d</v></c>%s%s</row>' % (r, st[0], lineno, c_ts, c_tx))

    def _runs(self, text, hits):
        out, pos, base = [], 0, self.rp_base
        for a, b, k in hits:
            if a > pos:
                out.append('<r>%s<t xml:space="preserve">%s</t></r>' % (base, _xs(text[pos:a])))
            out.append('<r>%s<t xml:space="preserve">%s</t></r>' % (self.rp_kw[k], _xs(text[a:b])))
            pos = b
        if pos < len(text):
            out.append('<r>%s<t xml:space="preserve">%s</t></r>' % (base, _xs(text[pos:])))
        return ''.join(out)

    def discard(self):
        for sh in self.sheets:
            self.book.discard_sheet(sh)
        self.sheets = []

    def finish(self):
        """読み込みの完了。集計結果を FileResult に反映する"""
        res = self.res
        for sh in self.sheets:
            sh.fh.close()
        res.sheet_names = [sh.name for sh in self.sheets]
        res.matched = sum(sh.data_rows for sh in self.sheets)
        res.kw_hits = [sum(sh.kw_hits[k] for sh in self.sheets) for k in range(self.n_kw)]

    def write(self):
        """シートを Excel ファイルに書き込む"""
        total = len(self.sheets)
        for i, sh in enumerate(self.sheets, 1):
            tab = next((KW_PALETTE[k % len(KW_PALETTE)][0] for k, c in enumerate(sh.kw_hits) if c), None)
            self.book.finalize_data_sheet(sh, self._header_rows(sh, i, total), tab)

    def _header_rows(self, sh, i, total):
        o, res, cfg, ds = self.obj, self.res, self.cfg, self.book.ds
        title = o.key.rstrip('/').rsplit('/', 1)[-1] + ('（%d/%d）' % (i, total) if total > 1 else '')
        fmt = encoding_label(res.encoding) + ('・%s' % res.compression if res.compression else '')
        runs2 = [('S3  ', RP_LABEL), ('s3://%s/%s' % (cfg.bucket, o.key), RP_VALUE), SEP,
                 ('最終更新  ', RP_LABEL), (fmt_dt(o.mtime, cfg.tz), RP_VALUE), SEP,
                 ('サイズ  ', RP_LABEL), (human_size(o.size), RP_VALUE), SEP,
                 ('形式  ', RP_LABEL), (fmt, RP_VALUE)]
        if res.ts_label:
            runs2 += [SEP, ('日時の形式  ', RP_LABEL), (res.ts_label, RP_VALUE)]
        runs3 = [('抽出  ', RP_LABEL), ('%s 行' % fmt_n(res.matched), RP_STRONG),
                 (' / 全 %s 行' % fmt_n(res.total_lines or 0), RP_VALUE)]
        if total > 1:
            runs3.append(('（このシート %s 行）' % fmt_n(sh.data_rows), RP_MUTED))
        runs3 += [SEP, ('期間  ', RP_LABEL), (cfg.range_text, RP_VALUE)]
        if sh.ts_min is not None:
            runs3 += [SEP, ('ログの時刻  ', RP_LABEL), ('%s ～ %s' % (
                fmt_dt(sh.ts_min, cfg.tz, True), fmt_dt(sh.ts_max, cfg.tz, True)), RP_VALUE)]
        if cfg.keywords:
            runs3 += [SEP, ('強調  ', RP_LABEL)]
            for k, kw in enumerate(cfg.keywords):
                c = sh.kw_hits[k]
                runs3 += [(kw, rpr(9, KW_PALETTE[k % len(KW_PALETTE)][0], bold=True)),
                          (' %s　' % fmt_n(c), RP_VALUE if c else RP_MUTED)]
        if res.untimed_all:
            runs3 += [SEP, ('※ 日時を検出できなかったため全行を出力しています', RP_NOTE)]
        return ''.join((
            row_xml(1, [c_str('A1', ds.title, title), c_blank('B1', ds.title_rule), c_blank('C1', ds.title_rule)],
                    ht=27),
            row_xml(2, [c_rich('A2', ds.meta, runs2)], ht=18),
            row_xml(3, [c_rich('A3', ds.meta, runs3)], ht=18),
            row_xml(4, [c_str('A4', ds.head_c, '行番号'), c_str('B4', ds.head_c, '日時 (%s)' % cfg.tz_short),
                        c_str('C4', ds.head_l, 'ログ')], ht=24),
        ))


def process_object(obj, path, cfg, book):
    """ダウンロードしたファイルを読み、期間内の行をシートに書き出す。
    (結果, シート) を返す。シートは呼び出し元で write() して Excel ファイルに書き込む。"""
    res = FileResult(obj)
    if obj.size == 0:
        res.status = 'empty'
        return res, None
    comp = sniff_compression(path)
    res.compression = comp
    if comp == 'zip':
        res.status, res.message = 'unsupported', 'zip'
        return res, None
    try:
        enc = cfg.encoding or detect_encoding(path, comp)
    except _READ_ERRORS as e:
        res.status, res.message = 'error', '展開できません（%s）' % (str(e) or e.__class__.__name__)
        return res, None
    if enc is None:
        res.status = 'binary'
        return res, None
    res.encoding = enc

    ctx = TsContext(cfg.local_offset, obj.mtime)
    if cfg.ts_custom is not None:
        parser, res.ts_label = cfg.ts_custom.bind(ctx), cfg.ts_custom.label
    else:
        gen = iter_lines(path, comp, enc, ReadState())
        try:
            sample = list(itertools.islice(gen, SAMPLE_LINES))
        finally:
            gen.close()
        parser, res.ts_label = detect_time_parser(sample, ctx)
    if parser is None:
        if cfg.untimed == 'skip' or (cfg.untimed == 'mtime' and not cfg.start <= obj.mtime < cfg.end):
            res.status = 'untimed' if cfg.untimed == 'skip' else 'none'
            return res, None
        res.untimed_all = True

    sheets = FileSheets(book, obj, cfg, res)
    state = ReadState()
    start, end = cfg.start, cfg.end
    cur = None
    band = 1
    lineno = 0
    try:
        for lineno, line in enumerate(iter_lines(path, comp, enc, state), 1):
            if parser is not None:
                t = parser(line)
                if t is None:                       # 日時のない行は直前の行と同じ日時として扱う
                    if cur is None or not start <= cur < end:
                        continue
                    cont = True
                else:
                    cur = t
                    if not start <= t < end:
                        continue
                    cont = False
                    band ^= 1
            else:
                cont = False
                band ^= 1
            sheets.add(lineno, cur, cont, line, band)
    except BaseException:
        sheets.discard()
        raise
    res.total_lines = lineno
    sheets.finish()
    if state.error:
        res.status, res.message = 'partial', '圧縮データが途中で壊れています'
    elif not res.matched:
        res.status = 'none'
    return res, sheets


# =============================================================================
# サマリーシート
# =============================================================================
def status_text(res):
    s = res.status
    if s == 'ok':
        if len(res.sheet_names) > 1:
            return '出力（%d シートに分割）' % len(res.sheet_names), 'ok'
        if res.untimed_all:
            return '出力（日時なし・全行）', 'warn'
        return '出力', 'ok'
    return {
        'partial': ('一部のみ出力（%s）' % res.message, 'warn'),
        'none': ('期間内のログなし', 'muted'),
        'untimed': ('日時を検出できず', 'warn'),
        'empty': ('空のファイル', 'muted'),
        'binary': ('テキストではないため対象外', 'muted'),
        'unsupported': ('未対応の形式（%s）' % res.message, 'warn'),
    }.get(s, ('エラー：%s' % res.message, 'error'))


def build_summary(book, cfg, results, stats):
    sb = book.sb
    kws = cfg.keywords
    nk = len(kws)
    ncol = 9 + nk                   # A:余白 B:No. C:シート D:キー E:最終更新 F:サイズ G:抽出 H:総行数 [強調…] 状態
    last = col_letter(ncol)
    kw_col = [col_letter(9 + k) for k in range(nk)]
    v = 'center'

    rule = sb.border(bottom=('medium', C_NAVY))
    s_title = sb.xf(font=sb.font(size=16, bold=True, color=C_NAVY), border=rule, v=v)
    s_rule = sb.xf(border=rule)
    s_label = sb.xf(font=sb.font(size=9, bold=True, color=C_SUB), v=v)
    s_value = sb.xf(font=sb.font(), v=v)
    hb = sb.border(left=('thin', C_NAVY_LINE), right=('thin', C_NAVY_LINE), top=('thin', C_NAVY),
                   bottom=('thin', C_NAVY))
    s_head = sb.xf(font=sb.font(bold=True, color=C_WHITE), fill=sb.fill(C_NAVY), border=hb, h='center', v=v)
    s_head_kw = [sb.xf(font=sb.font(bold=True, color=KW_PALETTE[k % len(KW_PALETTE)][0]),
                       fill=sb.fill(KW_PALETTE[k % len(KW_PALETTE)][1]), border=hb, h='center', v=v)
                 for k in range(nk)]
    grid = ('thin', C_GRID)

    def body_styles(fill_rgb, border):
        fi = sb.fill(fill_rgb)
        f_text, f_mut = sb.font(), sb.font(color=C_MUTED)
        return {
            'no': sb.xf(font=f_mut, fill=fi, border=border, h='center', v=v),
            'link': sb.xf(font=sb.font(color=C_LINK, underline=True), fill=fi, border=border, h='left', v=v,
                          indent=1),
            'nolink': sb.xf(font=f_mut, fill=fi, border=border, h='center', v=v),
            'text': sb.xf(font=f_text, fill=fi, border=border, h='left', v=v, indent=1),
            'dt': sb.xf(font=f_text, fill=fi, border=border, numfmt=FMT_DT, h='center', v=v),
            'size': sb.xf(font=f_text, fill=fi, border=border, numfmt=FMT_SIZE, h='right', v=v, indent=1),
            'int': sb.xf(font=sb.font(bold=True), fill=fi, border=border, numfmt=FMT_INT, h='right', v=v, indent=1),
            'int_m': sb.xf(font=f_mut, fill=fi, border=border, numfmt=FMT_INT, h='right', v=v, indent=1),
            'kw': [sb.xf(font=sb.font(bold=True, color=KW_PALETTE[k % len(KW_PALETTE)][0]), fill=fi, border=border,
                         numfmt=FMT_INT, h='center', v=v) for k in range(nk)],
            'kw0': sb.xf(font=f_mut, fill=fi, border=border, numfmt=FMT_INT, h='center', v=v),
            'status': {c: sb.xf(font=sb.font(color=rgb, bold=c in ('warn', 'error')), fill=fi, border=border,
                                h='left', v=v, indent=1) for c, rgb in STATUS_COLOR.items()},
            'blank': sb.xf(fill=fi, border=border, v=v),
            'total': sb.xf(font=sb.font(bold=True, color=C_NAVY), fill=fi, border=border, h='left', v=v, indent=1),
        }
    body_bd = sb.border(left=grid, right=grid, bottom=grid)
    bodies = (body_styles(None, body_bd), body_styles(C_BAND, body_bd))
    tot = body_styles(None, sb.border(top=('medium', C_NAVY), bottom=('thin', C_NAVY)))

    # ---- 見出し部分
    n_sheets = sum(len(r.sheet_names) for r in results)
    n_lines = sum(r.matched for r in results)
    counts = collections.Counter(r.status for r in results)
    detail = []
    if stats['skipped_mtime']:
        detail.append('最終更新日時が期間より前のため %d 件を除外' % stats['skipped_mtime'])
    for key, label in (('none', '期間内のログなし'), ('untimed', '日時を検出できず'), ('error', 'エラー'),
                       ('partial', '一部のみ出力'), ('binary', 'テキスト以外'), ('empty', '空のファイル'),
                       ('unsupported', '未対応の形式')):
        if counts.get(key):
            detail.append('%s %d 件' % (label, counts[key]))
    result_runs = [('%d ファイルを確認し、%d シートに %s 行を出力' % (len(results), n_sheets, fmt_n(n_lines)),
                    rpr(10, C_TEXT, bold=True))]
    if detail:
        result_runs.append(('　（%s）' % '、'.join(detail), rpr(9, C_SUB)))
    if kws:
        kw_runs = []
        for k, kw in enumerate(kws):
            kw_runs += [(kw, rpr(10, KW_PALETTE[k % len(KW_PALETTE)][0], bold=True)), ('　　', rpr(10))]
        opts = [o for o, on in (('大文字・小文字を区別しない', cfg.ignore_case), ('正規表現', cfg.regex)) if on]
        if opts:
            kw_runs.append(('（%s）' % '・'.join(opts), rpr(9, C_SUB)))
    else:
        kw_runs = [('（指定なし）', rpr(10, C_MUTED))]
    now = time.time()
    info = (
        ('対象 S3', [(cfg.s3uri, rpr(10))]),
        ('抽出期間', [(cfg.range_text, rpr(10)), ('　（%s）' % cfg.tz_long, rpr(9, C_SUB))]),
        ('強調文言', kw_runs),
        ('処理結果', result_runs),
        ('実行 ID', [(cfg.identity or '（不明）', rpr(10))]),
        ('作成日時', [(fmt_dt(now, cfg.tz), rpr(10)), ('　s3log2xlsx %s' % cfg.version, rpr(9, C_MUTED))]),
    )
    rows = [row_xml(1, [], ht=10),
            row_xml(2, [c_str('B2', s_title, 'S3 ログ抽出レポート')] +
                    [c_blank('%s2' % col_letter(c), s_rule) for c in range(3, ncol + 1)], ht=32),
            row_xml(3, [], ht=8)]
    for i, (label, runs) in enumerate(info):
        r = 4 + i
        rows.append(row_xml(r, [c_str('B%d' % r, s_label, label), c_rich('C%d' % r, s_value, runs)], ht=19))
    head_row = 4 + len(info) + 1
    rows.append(row_xml(head_row - 1, [], ht=12))
    heads = ['No.', 'シート', 'S3 キー', '最終更新', 'サイズ', '抽出行数', '総行数']
    cells = [c_str('%s%d' % (col_letter(2 + j), head_row), s_head, h) for j, h in enumerate(heads)]
    cells += [c_str('%s%d' % (kw_col[k], head_row), s_head_kw[k], kws[k]) for k in range(nk)]
    cells.append(c_str('%s%d' % (last, head_row), s_head, '状態'))
    rows.append(row_xml(head_row, cells, ht=24))

    # ---- ファイルごとの行
    links = []
    widths = {'sheet': 10, 'key': 16, 'status': 10}
    totals = {'size': 0, 'matched': 0, 'lines': 0, 'kw': [0] * nk}
    r = head_row
    for i, res in enumerate(results, 1):
        r += 1
        S = bodies[(i - 1) % 2]
        o = res.obj
        cells = [c_num('B%d' % r, S['no'], i)]
        if res.sheet_names:
            nm = res.sheet_names[0]
            cells.append(c_str('C%d' % r, S['link'], nm))
            links.append(('C%d' % r, '%s!A1' % quote_sheet(nm), nm))
            widths['sheet'] = max(widths['sheet'], text_width(nm))
        else:
            cells.append(c_str('C%d' % r, S['nolink'], '―'))
        cells.append(c_str('D%d' % r, S['text'], o.key))
        widths['key'] = max(widths['key'], text_width(o.key))
        cells.append(c_num('E%d' % r, S['dt'], cfg.display_offset.serial(o.mtime)))
        cells.append(c_num('F%d' % r, S['size'], o.size))
        totals['size'] += o.size
        if res.total_lines is None:
            cells += [c_blank('G%d' % r, S['blank']), c_blank('H%d' % r, S['blank'])]
        else:
            cells.append(c_num('G%d' % r, S['int'] if res.matched else S['int_m'], res.matched))
            cells.append(c_num('H%d' % r, S['int_m'], res.total_lines))
            totals['matched'] += res.matched
            totals['lines'] += res.total_lines
        for k in range(nk):
            if res.kw_hits is None:
                cells.append(c_blank('%s%d' % (kw_col[k], r), S['blank']))
            else:
                n = res.kw_hits[k]
                totals['kw'][k] += n
                cells.append(c_num('%s%d' % (kw_col[k], r), S['kw'][k] if n else S['kw0'], n))
        text, cat = status_text(res)
        widths['status'] = max(widths['status'], text_width(text))
        cells.append(c_str('%s%d' % (last, r), S['status'][cat], text))
        rows.append(row_xml(r, cells, ht=20))
    last_data = r
    r += 1
    cells = [c_str('B%d' % r, tot['total'], '合計'), c_blank('C%d' % r, tot['blank']), c_blank('D%d' % r, tot['blank']),
             c_blank('E%d' % r, tot['blank']), c_num('F%d' % r, tot['size'], totals['size']),
             c_num('G%d' % r, tot['int'], totals['matched']), c_num('H%d' % r, tot['int_m'], totals['lines'])]
    cells += [c_num('%s%d' % (kw_col[k], r), tot['kw'][k] if totals['kw'][k] else tot['kw0'], totals['kw'][k])
              for k in range(nk)]
    cells.append(c_blank('%s%d' % (last, r), tot['blank']))
    rows.append(row_xml(r, cells, ht=22))

    cols = [2.5, 10.0, min(36.0, widths['sheet'] + 4.0), min(70.0, widths['key'] + 3.0), 21.0, 11.5, 11.5, 11.5]
    cols += [min(24.0, max(10.0, text_width(kw) + 5.0)) for kw in kws]
    cols.append(min(44.0, widths['status'] + 4.0))
    filter_ref = '$B$%d:$%s$%d' % (head_row, last, last_data) if results else None
    xml = (sheet_head('A1:%s%d' % (last, r), cols, tab_color=C_NAVY, selected=True, fit_width=True) +
           ''.join(rows) +
           sheet_tail(autofilter=filter_ref.replace('$', '') if filter_ref else None, hyperlinks=links,
                      fit_width=True))
    book.write_summary(xml, filter_ref, head_row)


# =============================================================================
# aws CLI
# =============================================================================
_AWS_ERR_RE = re.compile(r'An error occurred \(([^)]+)\)(?: when calling the \w+ operation)?:\s*(.*)')


def aws_error_summary(err):
    lines = [l.strip() for l in err.splitlines() if l.strip()]
    for l in reversed(lines):
        m = _AWS_ERR_RE.search(l)
        if m:
            code, msg = m.group(1), m.group(2).strip()
            if code == 'InvalidObjectState':
                return 'アーカイブ（Glacier）済みのため取得できません。復元が必要です'
            return '%s: %s' % (code, msg) if msg else code
    if lines:
        return re.sub(r'^aws: \[ERROR\]:\s*', '', lines[-1])
    return '不明なエラー'


class AwsCli:
    def __init__(self, region, con):
        self.region = region
        self.con = con
        self.env = dict(os.environ, AWS_PAGER='', AWS_CLI_AUTO_PROMPT='off')

    def run(self, args):
        cmd = ['aws'] + list(args) + (['--region', self.region] if self.region else [])
        self.con.debug('$ ' + ' '.join(shlex.quote(c) for c in cmd))
        try:
            p = subprocess.run(cmd, stdin=subprocess.DEVNULL, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                               env=self.env)
        except OSError as e:
            raise FatalError('aws コマンドを実行できません: %s' % e)
        return p.returncode, p.stdout, p.stderr.decode('utf-8', 'replace')

    def list_objects(self, bucket, prefix):
        args = ['s3api', 'list-objects-v2', '--bucket=' + bucket]
        if prefix:
            args.append('--prefix=' + prefix)
        args += ['--output', 'json', '--query', 'Contents[].{Key:Key,Size:Size,LastModified:LastModified}']
        rc, out, err = self.run(args)
        if rc != 0:
            raise FatalError('S3 のオブジェクト一覧を取得できませんでした: %s' % aws_error_summary(err))
        text = out.decode('utf-8', 'replace').strip()
        try:
            data = json.loads(text) if text else None
        except ValueError as e:
            raise FatalError('S3 のオブジェクト一覧を解釈できませんでした: %s' % e)
        objs = []
        for item in data or []:
            objs.append(S3Object(item['Key'], int(item.get('Size') or 0), parse_s3_time(item['LastModified'])))
        return objs

    def get_object(self, bucket, key, dest):
        rc, _out, err = self.run(['s3api', 'get-object', '--bucket=' + bucket, '--key=' + key, dest])
        if rc != 0:
            try:
                os.remove(dest)
            except OSError:
                pass
            return aws_error_summary(err)
        return None


def parse_s3_time(s):
    s = s.strip()
    if s.endswith('Z'):
        s = s[:-1] + '+00:00'
    m = re.match(r'^(.*?\.\d{1,6})\d*([+-]\d{2}:\d{2})$', s)
    if m:
        s = m.group(1) + m.group(2)
    m = re.match(r'^(.*?)\.(\d{1,6})([+-].*)$', s)
    if m:
        s = '%s.%s%s' % (m.group(1), m.group(2).ljust(6, '0'), m.group(3))
    d = dt.datetime.fromisoformat(s)
    if d.tzinfo is None:
        d = d.replace(tzinfo=UTC)
    return d.timestamp()


class Downloader:
    """aws CLI を並列に実行して先読みし、取得した順ではなくキーの順に返す"""

    def __init__(self, aws, bucket, dldir, parallel):
        self.aws, self.bucket, self.dldir, self.parallel = aws, bucket, dldir, parallel
        self.pool = concurrent.futures.ThreadPoolExecutor(max_workers=parallel)

    def _fetch(self, idx, obj):
        path = os.path.join(self.dldir, 'obj%06d' % idx)
        err = self.aws.get_object(self.bucket, obj.key, path)
        return (None, err) if err else (path, None)

    def iterate(self, objs):
        queue = collections.deque()
        it = iter(enumerate(objs))

        def feed():
            for idx, obj in it:
                queue.append((obj, self.pool.submit(self._fetch, idx, obj)))
                return True
            return False
        for _ in range(self.parallel * 2):
            if not feed():
                break
        while queue:
            obj, fut = queue.popleft()
            feed()
            yield obj, fut.result()

    def close(self):
        self.pool.shutdown(wait=True, cancel_futures=True)
        for name in os.listdir(self.dldir):
            try:
                os.remove(os.path.join(self.dldir, name))
            except OSError:
                pass


# =============================================================================
# 設定と実行
# =============================================================================
class ArgParser(argparse.ArgumentParser):
    def error(self, message):
        raise UsageError('引数の誤り: %s' % message)


def build_arg_parser():
    p = ArgParser(prog='s3log2xlsx-core', add_help=False, allow_abbrev=False)
    p.add_argument('mode', choices=('check', 'run'))
    p.add_argument('--bucket', required=True)
    p.add_argument('--prefix', default='')
    p.add_argument('--from', dest='time_from', required=True)
    p.add_argument('--to', dest='time_to', default='now')
    p.add_argument('--now', type=float, default=0.0)
    p.add_argument('--tz', default='')
    p.add_argument('--log-tz', default='')
    p.add_argument('--output', required=True)
    p.add_argument('--keyword', action='append', default=[])
    p.add_argument('--ignore-case', action='store_true')
    p.add_argument('--regex', action='store_true')
    p.add_argument('--include', action='append', default=[])
    p.add_argument('--exclude', action='append', default=[])
    p.add_argument('--mtime-margin', type=int, default=60)
    p.add_argument('--max-files', type=int, default=500)
    p.add_argument('--parallel', type=int, default=4)
    p.add_argument('--encoding', default='')
    p.add_argument('--ts-regex', default='')
    p.add_argument('--ts-format', default='')
    p.add_argument('--untimed', choices=('skip', 'all', 'mtime'), default='skip')
    p.add_argument('--wrap', action='store_true')
    p.add_argument('--force', action='store_true')
    p.add_argument('--region', default='')
    p.add_argument('--identity', default='')
    p.add_argument('--workdir', default='')
    p.add_argument('--tool-version', default='')
    p.add_argument('--verbose', action='store_true')
    p.add_argument('--max-rows-per-sheet', type=int, default=0, help=argparse.SUPPRESS)
    return p


TIME_HINT = '書式の例: "2026-10-04 09:00"  "2026/10/04 09:00:30"  "2026-10-04T09:00:00+09:00"  "09:00"  -2h  now'


def normalize_output(path, force):
    p = os.path.abspath(os.path.expanduser(path))
    ext = os.path.splitext(p)[1]
    if not ext:
        p += '.xlsx'
    elif ext.lower() != '.xlsx':
        raise UsageError('出力ファイルの拡張子は .xlsx にしてください: %s' % path)
    d = os.path.dirname(p)
    if not os.path.isdir(d):
        raise UsageError('出力先のディレクトリがありません: %s' % d)
    if not os.access(d, os.W_OK):
        raise UsageError('出力先のディレクトリに書き込めません: %s' % d)
    if os.path.isdir(p):
        raise UsageError('出力先がディレクトリです: %s' % p)
    if os.path.exists(p) and not force:
        raise UsageError('出力ファイルが既にあります: %s' % p, '上書きする場合は -F/--force を指定してください。')
    return p


def build_config(a):
    cfg = argparse.Namespace()
    cfg.bucket, cfg.prefix = a.bucket, a.prefix
    cfg.s3uri = 's3://%s/%s' % (a.bucket, a.prefix)
    try:
        cfg.tz = parse_tz(a.tz)
    except ValueError:
        raise UsageError('--tz に指定したタイムゾーンが不明です: %s（例: Asia/Tokyo, UTC, +09:00）' % a.tz)
    try:
        cfg.log_tz = parse_tz(a.log_tz) if a.log_tz else cfg.tz
    except ValueError:
        raise UsageError('--log-tz に指定したタイムゾーンが不明です: %s（例: Asia/Tokyo, UTC, +09:00）' % a.log_tz)
    now = dt.datetime.fromtimestamp(a.now, UTC) if a.now else dt.datetime.now(UTC)
    try:
        start, _step = parse_user_time(a.time_from, cfg.tz, now)
    except ValueError:
        raise UsageError('--from の日時を解釈できません: %s' % a.time_from, TIME_HINT)
    try:
        to_start, step = parse_user_time(a.time_to, cfg.tz, now)
    except ValueError:
        raise UsageError('--to の日時を解釈できません: %s' % a.time_to, TIME_HINT)
    cfg.start, cfg.end = start.timestamp(), (to_start + step).timestamp()
    if cfg.start >= cfg.end:
        raise UsageError('抽出開始日時が終了日時より後になっています: %s ～ %s' % (
            fmt_dt(cfg.start, cfg.tz, True), fmt_dt(cfg.end - 1e-6, cfg.tz, True)))
    cfg.range_text = '%s ～ %s' % (fmt_dt(cfg.start, cfg.tz, True), fmt_dt(cfg.end - 1e-6, cfg.tz, True))
    cfg.tz_long = tz_long(cfg.tz, cfg.start)
    cfg.tz_short = tz_short(cfg.tz, cfg.start)

    cfg.keywords = list(collections.OrderedDict.fromkeys(k for k in a.keyword if k))
    cfg.ignore_case, cfg.regex = a.ignore_case, a.regex
    cfg.highlighter = Highlighter(cfg.keywords, a.ignore_case, a.regex) if cfg.keywords else None
    cfg.encoding = normalize_encoding(a.encoding)
    cfg.ts_custom = CustomTimeSpec(a.ts_regex, a.ts_format) if a.ts_regex else None
    if a.ts_format and not a.ts_regex:
        raise UsageError('--ts-format は --ts-regex と組み合わせて指定してください。')
    cfg.output = normalize_output(a.output, a.force)
    cfg.includes, cfg.excludes = a.include, a.exclude
    cfg.mtime_margin, cfg.max_files, cfg.untimed, cfg.wrap = a.mtime_margin, a.max_files, a.untimed, a.wrap
    cfg.parallel = min(16, max(1, a.parallel))
    cfg.region, cfg.identity, cfg.version = a.region, a.identity, a.tool_version
    cfg.workdir = a.workdir
    cfg.max_rows = a.max_rows_per_sheet or (EXCEL_MAX_ROWS - HEADER_ROWS)
    cfg.display_offset = DisplayOffset(cfg.tz)
    cfg.local_offset = LocalOffset(cfg.log_tz)
    return cfg


def cmd_check(cfg, con):
    con.kv('対象 S3', cfg.s3uri)
    con.kv('抽出期間', '%s（%s）' % (cfg.range_text, cfg.tz_long))
    if cfg.keywords:
        opts = [o for o, on in (('大文字・小文字を区別しない', cfg.ignore_case), ('正規表現', cfg.regex)) if on]
        con.kv('強調文言', ', '.join(cfg.keywords) + ('（%s）' % '・'.join(opts) if opts else ''))
    con.kv('出力先', cfg.output)
    return EXIT_OK


def select_targets(objs, cfg):
    stats = collections.Counter()
    margin = cfg.mtime_margin * 60
    targets = []
    for o in objs:
        if o.key.endswith('/'):
            stats['skipped_dir'] += 1
            continue
        if cfg.includes and not any(fnmatch.fnmatchcase(o.key, p) for p in cfg.includes):
            stats['skipped_pattern'] += 1
            continue
        if any(fnmatch.fnmatchcase(o.key, p) for p in cfg.excludes):
            stats['skipped_pattern'] += 1
            continue
        if cfg.mtime_margin >= 0 and o.mtime < cfg.start - margin:
            stats['skipped_mtime'] += 1
            continue
        targets.append(o)
    return targets, stats


def describe_result(res, kws):
    if res.status in ('ok', 'partial'):
        text = '%s 行を抽出' % fmt_n(res.matched)
        hits = ['%s %s' % (kw, fmt_n(n)) for kw, n in zip(kws, res.kw_hits or []) if n]
        if hits:
            text += '（%s）' % '・'.join(hits)
        if len(res.sheet_names) > 1:
            text += ' → %d シートに分割' % len(res.sheet_names)
        if res.status == 'partial':
            return text + '（%s）' % res.message, 'warn'
        return text, 'info'
    text, cat = status_text(res)
    if res.status == 'untimed':
        text += '（--ts-regex で日時の形式を指定するか、--untimed all で全行を出力できます）'
    return text, {'warn': 'warn', 'error': 'error'}.get(cat, 'info')


def cmd_run(cfg, con):
    t0 = time.time()
    aws = AwsCli(cfg.region, con)
    con.info('S3 のオブジェクト一覧を取得しています...')
    objs = aws.list_objects(cfg.bucket, cfg.prefix)
    if not objs:
        con.warn('%s にオブジェクトがありません。Excel は作成しません。' % cfg.s3uri)
        return EXIT_NO_TARGET
    targets, stats = select_targets(objs, cfg)
    skipped = []
    if stats['skipped_mtime']:
        skipped.append('最終更新日時が期間より前 %d 件' % stats['skipped_mtime'])
    if stats['skipped_pattern']:
        skipped.append('--include/--exclude による除外 %d 件' % stats['skipped_pattern'])
    if stats['skipped_dir']:
        skipped.append('フォルダー %d 件' % stats['skipped_dir'])
    con.info('%d 件中 %d ファイルを読み込みます%s' % (len(objs), len(targets),
                                                   '（除外: %s）' % '、'.join(skipped) if skipped else ''))
    if not targets:
        con.warn('条件に合うファイルがありません。Excel は作成しません。')
        return EXIT_NO_TARGET
    if len(targets) > cfg.max_files:
        raise FatalError('読み込むファイルが %d 件あり、上限（--max-files %d）を超えています。' % (
            len(targets), cfg.max_files),
            '-s のプレフィックスや --include / --exclude で対象を絞り込むか、--max-files で上限を引き上げてください。')

    if not cfg.workdir:
        cfg.workdir = tempfile.mkdtemp(prefix='s3log2xlsx.')
    dldir = os.path.join(cfg.workdir, 'download')
    os.makedirs(dldir, exist_ok=True)
    sb = StyleBook()
    book = XlsxBook(cfg.output, cfg.workdir, sb, cfg.wrap)
    book.ds = DataStyles(sb, cfg.wrap, len(cfg.keywords))
    results = []
    downloader = Downloader(aws, cfg.bucket, dldir, cfg.parallel)
    prefix = cfg.prefix if cfg.prefix.endswith('/') else ''
    try:
        try:
            for i, (obj, (path, err)) in enumerate(downloader.iterate(targets), 1):
                sheets = None
                if err:
                    res = FileResult(obj)
                    res.status, res.message = 'error', err
                else:
                    try:
                        res, sheets = process_object(obj, path, cfg, book)
                    except Exception as e:     # 1 ファイルの読み込み失敗で全体を止めない
                        res = FileResult(obj)
                        res.status, res.message = 'error', '読み込みに失敗しました（%s: %s）' % (
                            e.__class__.__name__, e)
                    finally:
                        try:
                            os.remove(path)
                        except OSError:
                            pass
                if sheets is not None:
                    sheets.write()             # Excel ファイルへの書き込みの失敗は致命的エラーとする
                results.append(res)
                name = obj.key[len(prefix):] if prefix and obj.key.startswith(prefix) else obj.key
                text, level = describe_result(res, cfg.keywords)
                con.progress(i, len(targets), '%s (%s) … %s' % (name, human_size(obj.size), text), level)
        finally:
            downloader.close()
        build_summary(book, cfg, results, stats)
        book.close('S3 ログ抽出レポート（%s）' % cfg.s3uri, dt.datetime.now(UTC))
    except BaseException:
        book.abort()
        raise

    n_sheets = sum(len(r.sheet_names) for r in results)
    n_lines = sum(r.matched for r in results)
    errors = sum(1 for r in results if r.status == 'error')
    partial = sum(1 for r in results if r.status == 'partial')
    con.ok('Excel ファイルを作成しました: %s' % cfg.output,
           'シート %d 枚（サマリー + %d）、抽出 %s 行、処理時間 %.1f 秒' % (
               n_sheets + 1, n_sheets, fmt_n(n_lines), time.time() - t0))
    if not n_lines:
        con.warn('指定期間のログは見つかりませんでした（サマリーのみ出力しています）。')
    if errors or partial:
        con.warn('読み込めなかったファイルが %d 件、一部しか読み込めなかったファイルが %d 件あります。' % (
            errors, partial), 'サマリーシートの「状態」を確認してください。')
        return EXIT_PARTIAL
    return EXIT_OK


def _on_sigterm(_signum, _frame):
    raise Terminated()


def main(argv):
    con = Console()
    signal.signal(signal.SIGTERM, _on_sigterm)
    try:
        a = build_arg_parser().parse_args(argv)
        con.verbose = a.verbose
        cfg = build_config(a)
        return cmd_check(cfg, con) if a.mode == 'check' else cmd_run(cfg, con)
    except UsageError as e:
        con.error(*e.args)
        return EXIT_USAGE
    except FatalError as e:
        con.error(*e.args)
        return EXIT_ERROR
    except KeyboardInterrupt:
        con.error('中断しました。')
        return 130
    except Terminated:
        con.error('終了要求を受けたため中断しました。')
        return 143


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
S3X_PY_EOF
}

s3x_main "$@"
