#!/bin/zsh
set -u
APP_DIR="$(cd "$(dirname "$0")" && pwd)"
NODE_BIN="$APP_DIR/runtime/node/bin/node"
if [[ ! -x "$NODE_BIN" ]]; then NODE_BIN="$(command -v node)"; fi
if [[ ! -x "$NODE_BIN" ]]; then
  print "缺少自带运行环境：$NODE_BIN"
  read "?按回车退出。"
  exit 1
fi
"$NODE_BIN" "$APP_DIR/app/launch.mjs" --preview
RUN_STATUS=$?
if [[ $RUN_STATUS -eq 0 ]]; then
  print "运行结束。结果保存在 output 中本次表格对应的文件夹。"
else
  print "运行未正常完成，请查看上方具体错误。"
fi
read "?按回车关闭窗口。"
exit $RUN_STATUS
