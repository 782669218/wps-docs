**OLEFormat**



包含 OLE 对象的属性。

**说明**

如果 **Shape** 对象不代表某个链接的或嵌入的对象，则 **OLEFormat** 属性失败。

**方法**

|                                                              | 名称         | 说明                              |
| ------------------------------------------------------------ | ------------ | --------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Activate** | 使当前图表成为活动图表。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Verb**     | 向指定的 OLE 对象服务器发送动词。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Object**      | 返回与此 OLE 对象相联系的 OLE 自动化对象。**Object** 型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **progID**      | 返回对象的程序标识符。**String** 型，只读。                  |

**成员方法**

#### **OLEFormat.Activate**

使当前图表成为活动图表。

**语法**

**express.Activate()**

*express*   一个代表 **OLEFormat** 对象的变量。

#### **OLEFormat.Verb**

向指定的 OLE 对象服务器发送动词。

**语法**

**express.Verb(Verb)**

*express*   一个代表 **OLEFormat** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**  | **说明**                                                     |
| -------- | ------------- | ------------- | ------------------------------------------------------------ |
| *Verb*   | 可选          | **XlOLEVerb** | OLE 对象服务器将执行其操作的动词。如果省略此参数，则发送默认动词。对象的源应用程序决定哪些动词可用。OLE 对象的典型动词为 Open 和 Primary（由 XlOLEVerb 常量 xlOpen 和 xlPrimary 表示）。 |

**成员属性**

#### **OLEFormat.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **OLEFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test(){ 　　　　let myObject = ActiveWorkbook     if (myObject.Application.Value == "ET"){ 　　　　    MsgBox ("This is an ET Application object.") 　　　　} 　　　　else{} 　　　　    MsgBox ("This is not an ET Application object.") 　　　　}}` |

#### **OLEFormat.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **OLEFormat** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **OLEFormat.Object**

返回与此 OLE 对象相联系的 OLE 自动化对象。**Object** 型，只读。

**语法**

**express.Object**

*express*   一个代表 **OLEFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例在 sheet1 的内嵌 WPS 文档对象的开始处插入文字。注意，在 With 控制结构内的三条语句为 WordBasic 语句。*/ function test(){ 　　　　let wordObj = Worksheets.Item("Sheet1").OLEObjects(1) 　　　　wordObj.Activate() 　　　　let wbc = wordObj.Object.Application.WordBasic 　　　　wbc.StartOfDocument 　　　　wbc.Insert ("This is the beginning") 　　　　wbc.InsertPara }` |

#### **OLEFormat.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **OLEFormat** 对象的变量。

**示例**

#### **OLEFormat.progID**

返回对象的程序标识符。**String** 型，只读。

**语法**

**express.progID**

*express*   一个代表 **OLEFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例为第一张工作表中所有 OLE 对象创建程序标识符列表。*/ function test(){ 　　　　let rw = 0 　　　　for (let o = 1; o <= Worksheets.Item(1).OLEObjects.Count; o++){   　　　　  rw = rw + 1  　　　　   Worksheets.Item(2).cells.Item(rw, 1).Value2 = Worksheets.Item(1).OLEObjects(o).ProgId 　　　　} }` |

适用环境：web

适用平台：windows/linux