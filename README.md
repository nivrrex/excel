# About excel ole

---
win32 excel ole implementation for golang

在Windows系统下调用github.com/mattn/go-ole库操作Excel文件，需要go-ole库的支持

---
## 需求

``` bash
go get github.com/mattn/go-ole
go install github.com/mattn/go-ole
go get github.com/mattn/go-ole/oleutil
go install github.com/mattn/go-ole/oleutil
```

---
## 更新
2014.6.15 将原有函数调用模式，更新为struct + func 的调用模式，感觉面向对象一点，看起来稍显舒服。
2012.9.25 first commit.