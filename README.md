# 介绍

本程序有两部分组成

第一部分是VBA脚本，置入outlook的rules内执行。
负责把邮件中的附件（html类型的文件）保存到指定的文件夹。

第二部分是nodejs脚本，单独运行。
负责持续监测文件夹内是否有新增的html文件，如果有，则打开该文件，并且通过固定的操作流程来下载对应的pdf文件。


# 运行前提

* 需要安装nodejs
* 需要安装Chrome
* 需要确保outlook的rules内有'run a script'选项


# 注意事项

* 机器上的outlook和Chrome不要关闭，保持运行
* 上述nodejs脚本要确保一直在运行，不要退出
* chrome的setting里要把pdf设置为默认下载，而不是默认打开查看
* 可以修改Chrome的默认下载路径，从而指定pdf文件下载时保存到哪个文件夹

# 使用方法

## 1 将vba脚本导入outlook内

此时可以调整脚本内保存附件的文件夹路径

## 2 配置outlook rule

配置rule，在接收到特定邮件时，调用上述脚本来保存附件到特定文件夹

## 3 修改nodejs脚本的配置信息

找到index.js文件，并修改如下的配置信息

需要将配置设置为实际中使用的文件夹路径和密码，下面只是举例

```
//监控这个文件夹内新出现的html文件（也就是从邮件中保存的附件），需要设置的和outlook中保存附件的文件夹一致
const monitoredFolder = "C:/Data/Archive/";

//chrome 的 user profile文件夹  具体文件夹路径可以通过打开 chrome://version来查看
//需要先把chrome 设置为pdf默认下载而不是默认在页面内打开，否则无法下载，
const chromeProfilePath = "C:\\Users\\cpf\\AppData\\Local\\Google\\Chrome\\User Data";

// 用来下载pdf的密码，需替换
const pwd = '请替换此处的密码';
```

## 4 启动nodejs脚本