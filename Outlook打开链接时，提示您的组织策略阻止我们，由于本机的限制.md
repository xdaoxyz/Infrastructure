OUTLOOK2010提示：由于本机的限制，该操作已被取消。请与管理员联系。

OUTLOOK2013提示：您的组织策略阻止我们为您完成此操作。有关详细信息，请联系技术支持。

网上所说的注册表操作在域用户环境下并不是特别有用的

HKEY_CURRENT_USER\Software\Classes\.html 改成 Htmlfile
有时改掉了，重启了，依旧报错

解决办法为进入CMD，先查询当前数据数值

reg query HKEY_CURRENT_USER\Software\Classes\.html /ve
更改数值为Htmlfile

reg add HKEY_CURRENT_USER\Software\Classes\.html /ve /d Htmlfile
询问是否覆盖 输入yes即可

同样适用于WORD EXCEL 插入的超链接打不开的错误
