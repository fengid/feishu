## 快速开始 & demo
Forked from fastwego/feishu

```shell script
go get github.com/fengid/feishu
```
```go


// 创建 飞书 客户端
Client := feishu.NewClient()

// 调用 api 接口
var param UsersFindByDepartmentParam
data, err := Client.UsersFindByDepartment(param)
```