**CustomProperty**



代表标识符信息。标识符信息可用于 XML 的元数据。

**说明**

使用 **Add** 方法或 **CustomProperties** 集合的 **Item** 属性可返回 **CustomProperty** 对象。

返回 **CustomProperty** 对象后，可在 **Add** 方法中使用 **CustomProperties** 属性向工作表中添加元数据。

在本示例中，ET 向活动工作表添加标识符信息，并向用户返回名称和值。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let wksSheet1 = Application.ActiveSheet      // Add metadata to worksheet.     wksSheet1.CustomProperties.Add("Market", "Nasdaq")      // Display metadata.     let cusProperties = wksSheet1.CustomProperties.Item(1)     alert(cusProperties.Name + "\t" + cusProperties.Value) }` |

**方法**

|                                                              | 名称       | 说明       |
| ------------------------------------------------------------ | ---------- | ---------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete** | 删除对象。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**        | 返回或设置一个 **String** 值，它代表对象的名称。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Value**       | **Borders.LineStyle** 的同义词。                             |

**成员方法**

#### **CustomProperty.Delete**

删除对象。

**语法**

**express.Delete()**

*express*   一个代表 **CustomProperty** 对象的变量。

**说明**

可删除自定义文档属性，但是无法删除内置文档属性。

**成员属性**

#### **CustomProperty.Application**

如果不使用对象识别符，则该属性返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **CustomProperty** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例显示一条有关创建 myObject 的应用程序的消息*/ function test(){ 	let myObject = Application.ActiveWorkbook 	if(myObject.Application.Value == "ET") { 		alert("This is an ET Application object.") 	} 	else { 		alert("This is not an ET Application object.") 	} }` |

#### **CustomProperty.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **CustomProperty** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **CustomProperty.Name**

返回或设置一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

*express*   一个代表 **CustomProperty** 对象的变量。

#### **CustomProperty.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **CustomProperty** 对象的变量。

#### **CustomProperty.Value**

**Borders.LineStyle** 的同义词。

**语法**

**express.Value**

*express*   一个代表 **CustomProperty** 对象的变量。

适用环境：web

适用平台：windows/linux