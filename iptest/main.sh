#!/bin/bash

set -e          # 当任何命令以非零状态退出时，立即退出脚本
set -o pipefail # 确保管道中的每个命令都成功

# 定义日志函数
log() {
    echo "$(date +'%Y-%m-%d %H:%M:%S') - $1"
}

# 执行第一个 Node.js 脚本
log "开始执行 extraction.js"
if node extraction.js; then
    log "extraction.js 执行成功"
else
    log "extraction.js 执行失败"
    exit 1
fi

# 执行 Go 脚本
log "开始执行 iptest.go"
if go run iptest.go -file extraction_ip.txt -outfile detected_ip.csv -max 50; then
    log "iptest.go 执行成功"
else
    log "iptest.go 执行失败"
    exit 1
fi

# 执行第二个 Node.js 脚本
log "开始执行 detected.js"
if node detected.js; then
    log "detected.js 执行成功"
else
    log "detected.js 执行失败"
    exit 1
fi

log "所有脚本执行完毕"
