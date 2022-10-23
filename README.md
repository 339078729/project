# project
#golang配置
#强制开启GO111MODULE
	go env -w GO111MODULE=on
#使用中国代理
	go env -w GOPROXY=https://goproxy.cn,direct
#初始化mod
	go mod init 项目