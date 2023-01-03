**RecentFile**



代表最近使用的文件列表中的某个文件。

**说明**

**RecentFile** 对象是 **RecentFiles** 集合的成员

**方法**

|                                                              | 名称       | 说明                       |
| ------------------------------------------------------------ | ---------- | -------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete** | 删除对象。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Open**   | 打开一个最近使用的工作簿。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Index**       | 返回 **Long** 值，它代表对象在其同类对象所组成的集合内的索引号。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**        | 返回一个 **String** 值，它代表对象的名称。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Path**        | 返回一个 **String** 值，它代表应用程序的完整路径，不包括末尾的分隔符和应用程序名称。 |

**成员方法**

#### **RecentFile.Delete**

删除对象。

**语法**

**express.Delete()**

*express*   一个代表 **RecentFile** 对象的变量。

#### **RecentFile.Open**

打开一个最近使用的工作簿。

**语法**

**express.Open()**

*express*   一个代表 **RecentFile** 对象的变量。

**返回值**

一个代表打开的工作簿的 Workbook 对象。

**示例**

本示例打开 Analysis.xls 工作簿，然后运行它的 Auto_Open 宏。

| 示例代码复制                             |
| ---------------------------------------- |
| `Application.RecentFiles.Item(2).Open()` |

**成员属性**

#### **RecentFile.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **RecentFile** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test(){ 　　　　let myObject = ActiveWorkbook     if(myObject.Application.Value == "ET") { 　　　　    MsgBox("This is an ET Application object.") 　　　　} 　　　　else { 　　　　    MsgBox("This is not an ET Application object.") 　　　　} }` |

#### **RecentFile.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **RecentFile** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **RecentFile.Index**

返回 **Long** 值，它代表对象在其同类对象所组成的集合内的索引号。

**语法**

**express.Index**

*express*   一个代表 **RecentFile** 对象的变量。

#### **RecentFile.Name**

返回一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

*express*   一个代表 **RecentFile** 对象的变量。

#### **RecentFile.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **RecentFile** 对象的变量。

#### **RecentFile.Path**

返回一个 **String** 值，它代表应用程序的完整路径，不包括末尾的分隔符和应用程序名称。

**语法**

**express.Path**

*express*   一个代表 **RecentFile** 对象的变量。

适用环境：web

适用平台：windows/linux