这个Agent.py代码是因上一个项目写的一个客户端的功能，主要实现了：Config读写，Log日志 、 MySQL数据库操作 、 PYthon启动QTP 、 FTP客户端和ZIP打包，main（）主要是定时查询MYSQL当得到本机IP对应的任务时，根据计划执行时间得出时间戳。当到达执行时间时调有RunScenario()启动QTP。QTP执行完任务时将指定路径下的Report文件夹压缩成ZIP文件使用FTP上传到服务器。
该Agent在启动会一直会保持运行状态，来保证第一时间可以执行有最新的测试任务

refer:http://my.oschina.net/zhangzhe/blog/88042
