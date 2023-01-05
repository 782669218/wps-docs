**RoutingSlip**



代表工作簿的传送名单。传送名单用于在电子邮件系统中发送工作簿。

**说明**

**RoutingSlip** 对象不存在，无法返回该对象，除非工作簿的 **HasRoutingSlip** 属性是 **True**。

**方法**

|                                                              | 名称      | 说明                                                         |
| ------------------------------------------------------------ | --------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Reset** | 重新设置传送名单，从而可用同一名单初始化一个新的传送过程（使用同一收件人列表和收件人地址信息）。使用本方法之前，该传送过程必须已完成。否则本方法将导致错误。 |

**属性**

|                                                              | 名称               | 说明                                                         |
| ------------------------------------------------------------ | ------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**    | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**        | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Delivery**       | 返回或设置邮件传递的方法。可以是 **XlRoutingSlipDelivery** 常量之一。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Message**        | 返回或设置传送名单的消息文字。这些文字将作为传送指定工作簿时的邮件消息内容。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**         | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Recipients**     | 返回或设置传送名单上的收件人。                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ReturnWhenDone** | 如果传送过程结束后，工作簿返回给发件人，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Status**         | 指示传送名单的状态。**XlRoutingSlipStatus** 类型，只读。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Subject**        | 返回或设置邮件或传送名单的主题。**String** 型，可读写。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TrackStatus**    | 如果允许对传送名单进行状态追踪，则该属性值为 **True**。**Boolean** 类型，可读写。 |

**成员方法**

#### **RoutingSlip.Reset**

重新设置传送名单，从而可用同一名单初始化一个新的传送过程（使用同一收件人列表和收件人地址信息）。使用本方法之前，该传送过程必须已完成。否则本方法将导致错误。

**语法**

**express.Reset()**

*express*   一个代表 **RoutingSlip** 对象的变量。

**返回值**

Variant

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果工作簿 Book1.xls 的传送过程已完成，本示例将重新设置传送名单。*/ function test(){ 　　　　let slip= Workbooks.Item("BOOK1.XLS").RoutingSlip        if(slip.Status == xlRoutingComplete) { 　　　　    slip.Reset() 　　　　} 　　　　else { 　　　　    MsgBox("Cannot reset routing; not yet complete") 　　　　} }` |

**成员属性**

#### **RoutingSlip.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **RoutingSlip** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test(){ 　　　　let myObject = ActiveWorkbook        if(myObject.Application.Value == "ET") {  　　　　   MsgBox("This is an ET Application object.") 　　　　} 　　　　else{ 　　　　    MsgBox("This is not an ET Application object.") 　　　　}　　　　 }` |

#### **RoutingSlip.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **RoutingSlip** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **RoutingSlip.Delivery**

返回或设置邮件传递的方法。可以是 **XlRoutingSlipDelivery** 常量之一。**Long** 类型，可读写。

**语法**

**express.Delivery**

*express*   一个代表 **RoutingSlip** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 Book1.xls 依次发送给三个收件人。*/ function test(){ 　　　　Workbooks.Item("BOOK1.XLS").HasRoutingSlip = true 　　　　let slip= Workbooks.Item("BOOK1.XLS").RoutingSlip 　　　　slip.Delivery = xlOneAfterAnother 　　　　slip.Recipients = ["Adam Bendel","Jean Selva", "Bernard Gabor"] 　　　　slip.Subject = "Here is BOOK1.XLS" 　　　　slip.Message = "Here is the workbook. What do you think?" 　　　　Workbooks.Item("BOOK1.XLS").Route() }` |

#### **RoutingSlip.Message**

返回或设置传送名单的消息文字。这些文字将作为传送指定工作簿时的邮件消息内容。**String** 类型，可读写。

**语法**

**express.Message**

*express*   一个代表 **RoutingSlip** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 Book1.xls 依次发送给三个收件人。*/ function test(){ 　　　　Workbooks.Item("BOOK1.XLS").HasRoutingSlip = true 　　　　let slip= Workbooks.Item("BOOK1.XLS").RoutingSlip 　　　　slip.Delivery = xlOneAfterAnother 　　　　slip.Recipients = ["Adam Bendel","Jean Selva","Bernard Gabor"] 　　　　slip.Subject = "Here is BOOK1.XLS" 　　　　slip.Message = "Here is the workbook. What do you think?" 　　　　Workbooks.Item("BOOK1.XLS").Route() }` |

#### **RoutingSlip.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **RoutingSlip** 对象的变量。

#### **RoutingSlip.Recipients**

返回或设置传送名单上的收件人。

**语法**

**express.Recipients**

*express*   一个代表 **RoutingSlip** 对象的变量。

**说明**

如果传送过程的递送选项设置为 **xlOneAfterAnother**，则收件人列表的顺序就定义了传递的顺序。

如果某个传送名单正在进行中，则只会返回或设置那些还没有收到和传送该文档的收件人。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将打开的工作簿顺次发送给三个收件人。*/ function test(){ 　　　　let book= ThisWorkbook 　　　　book.HasRoutingSlip = true 　　　　let slip= book.RoutingSlip 　　　　slip.Delivery = xlOneAfterAnother 　　　　slip.Recipients = ["Adam Bendel","Jean Selva","Bernard Gabor"] 　　　　slip.Subject = "Here is the workbook" 　　　　slip.Message = "Here is the workbook. What do you think?" 　　　　slip.ReturnWhenDone = true 　　　　book.Route() }` |

#### **RoutingSlip.ReturnWhenDone**

如果传送过程结束后，工作簿返回给发件人，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ReturnWhenDone**

*express*   一个代表 **RoutingSlip** 对象的变量。

**说明**

如果正在进行传送，则无法设置该属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将工作簿 Book1.xls 按次序发送给三个收件人，传送过程结束后，再将工作簿返回给发件人。*/ function test(){ 　　　　Workbooks.Item("BOOK1.XLS").HasRoutingSlip = true 　　　　let slip= Workbooks.Item("BOOK1.XLS").RoutingSlip 　　　　slip.Delivery = xlOneAfterAnother 　　　　slip.Recipients = ["Adam Bendel","Jean Selva","Bernard Gabor"] 　　　　slip.Subject = "Here is BOOK1.XLS" 　　　　slip.Message = "Here is the workbook. What do you think?" 　　　　slip.ReturnWhenDone = true 　　　　Workbooks.Item("BOOK1.XLS").Route() }` |

#### **RoutingSlip.Status**

指示传送名单的状态。**XlRoutingSlipStatus** 类型，只读。

**语法**

**express.Status**

*express*   一个代表 **RoutingSlip** 对象的变量。

**说明**

| XlRoutingSlipStatus 可为以下 XlRoutingSlipStatus 常量之一。 |
| ----------------------------------------------------------- |
| xlNotYetRouted                                              |
| xlRoutingComplete                                           |
| xlRoutingInProgress                                         |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果工作簿 Book1.xls 的传送过程已完成，本示例将重新设置传送名单。*/ function test(){ 　　　　let slip= Workbooks.Item("BOOK1.XLS").RoutingSlip        if(slip.Status == xlRoutingComplete) {    　　　　 slip.Reset() 　　　　} 　　　　else {  　　　　   MsgBox("Cannot reset routing; not yet complete.") 　　　　} }` |

#### **RoutingSlip.Subject**

返回或设置邮件或传送名单的主题。**String** 型，可读写。

**语法**

**express.Subject**

*express*   一个代表 **RoutingSlip** 对象的变量。

**说明**

**RoutingSlip** 对象的主题被用作传送工作簿时的邮件信息的主题。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例设置打开的工作簿中传送名单的主题。此示例只能在已安装了 Microsoft Exchange 的情况下才能运行。*/ function test(){ 　　　　let book= ThisWorkbook 　　　　book.HasRoutingSlip = true 　　　　let slip=book.RoutingSlip 　　　　slip.Delivery = xlOneAfterAnother 　　　　slip.Recipients = ["Adam Bendel","Jean Selva","Bernard Gabor"] 　　　　slip.Subject = "Here is the workbook" 　　　　slip.Message = "Here is the workbook. What do you think?" 　　　　slip.ReturnWhenDone = true 　　　　book.Route() }` |

#### **RoutingSlip.TrackStatus**

如果允许对传送名单进行状态追踪，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.TrackStatus**

*express*   一个代表 **RoutingSlip** 对象的变量。

**说明**

如果正在进行传送，则无法设置该属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将工作簿 Book1.xls 发送给三个收件人，并允许进行状态追踪。*/ function test(){ 　　　　Workbooks.Item("BOOK1.XLS").HasRoutingSlip = true 　　　　let slip= Workbooks.Item("BOOK1.XLS").RoutingSlip 　　　　slip.Delivery = xlOneAfterAnother 　　　　slip.Recipients = ["Adam Bendel","Jean Selva", "Bernard Gabor"] 　　　　slip.Subject = "Here is BOOK1.XLS" 　　　　slip.Message = "Here is the workbook. What do you think?" 　　　　slip.ReturnWhenDone = true 　　　　slip.TrackStatus = true 　　　　Workbooks.Item("BOOK1.XLS").Route() }` |

适用环境：web

适用平台：windows/linux