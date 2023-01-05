**LinkFormat**



包含链接 OLE 对象的属性。

**说明**

如果指定 **Shape** 对象不代表一个链接对象，则 **LinkFormat** 属性会失败。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*使用 LinkFormat 属性可返回 LinkFormat 对象。下例更新 Shapes 集合中的一个 OLE 对象。*/ Application.Worksheets.Item(1).Shapes.Item(1).LinkFormat.Update()` |

**方法**

|                                                              | 名称       | 说明       |
| ------------------------------------------------------------ | ---------- | ---------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Update** | 更新链接。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoUpdate**  | 如果源更改时 **LinkFormat** 对象自动进行更新，则此属性为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Locked**      | 返回或设置一个 **Boolean** 值，它指明对象是否已被锁定。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |

**成员方法**

#### **LinkFormat.Update**

更新链接。

**语法**

**express.Update()**

*express*   一个代表 **LinkFormat** 对象的变量。

**成员属性**

#### **LinkFormat.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **LinkFormat** 对象的变量。

**说明**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() { 	let myObject = Application.ActiveWorkbook 	if (myObject.Application.Value == "ET") { 		alert("This is an ET Application object.") 	} else { 		alert("This is not an ET Application object.") 	} }` |

#### **LinkFormat.AutoUpdate**

如果源更改时 **LinkFormat** 对象自动进行更新，则此属性为 **True**。**Boolean** 类型，只读。

**语法**

**express.AutoUpdate**

*express*   一个代表 **LinkFormat** 对象的变量。

#### **LinkFormat.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **LinkFormat** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **LinkFormat.Locked**

返回或设置一个 **Boolean** 值，它指明对象是否已被锁定。

**语法**

**express.Locked**

*express*   一个代表 **LinkFormat** 对象的变量。

**说明**

如果对象已被锁定，此属性将返回 **True**；如果在工作表处于受保护状态时仍能修改对象，则返回 **False**。

#### **LinkFormat.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **LinkFormat** 对象的变量。

适用环境：web

适用平台：windows/linux