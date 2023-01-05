**TextFrame**



代表 **Shape** 对象中的文本框架。包含文本框架中的文本以及控制文本框架的对齐和定位的属性和方法。

**说明**

使用 **TextFrame** 属性可返回一个 **TextFrame** 对象。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例向 myDocument 中添加一个矩形，向矩形中添加文本，然后设置文本框的边距。*/ function test() {     let myDocument = Application.Worksheets.Item(1)     let rng = myDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 250,140).TextFrame      rng.Characters().Text = "Here is some test text"     rng.MarginBottom = 10     rng.MarginLeft = 10     rng.MarginRight = 10     rng.MarginTop = 10 }` |

**方法**

|                                                              | 名称           | 说明 |
| ------------------------------------------------------------ | -------------- | ---- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **DeleteText** |      |

**属性**

|                                                              | 名称                 | 说明                                                         |
| ------------------------------------------------------------ | -------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**      | 如果不使用对象识别符，则该属性返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoSize**         | 如果指定对象能自动调整大小，以适应其中所包含的文字，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HorizontalAnchor** |                                                              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MarginBottom**     | 以磅为单位返回或设置从文本框底端到包含文本的形状中内接矩形底端的距离。可读/写。**Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MarginLeft**       | 以磅为单位返回或设置从文本框左边界到包含文本的形状中内接矩形左边界的距离。可读写。**Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MarginRight**      | 以磅为单位返回或设置从文本框右边界到包含文本的形状中内接矩形右边界的距离。可读写。**Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MarginTop**        | 以磅为单位返回或设置从文本框架顶端到包含文本的形状中内接矩形顶端的距离。可读写。**Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Orientation**      | 返回或设置一个 **Long** 值，它代表文本框的方向。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Ruler**            |                                                              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextRange**        |                                                              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **VerticalAnchor**   |                                                              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WordWrap**         |                                                              |

**成员方法**

#### **TextFrame.DeleteText**

**语法**

**express.DeleteText()**

*express*   一个代表 **TextFrame** 对象的变量。

**成员属性**

#### **TextFrame.Application**

如果不使用对象识别符，则该属性返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。只读。

**语法**

**express.Application**

*express*   一个代表 **TextFrame** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() {   let myObject = Application.ActiveWorkbook   if(myObject.Application.Value == "ET"){       alert("This is an ET Application object.")   }   else{       alert("This is not an ET Application object.")   } }` |

#### **TextFrame.AutoSize**

如果指定对象能自动调整大小，以适应其中所包含的文字，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.AutoSize**

*express*   一个代表 **TextFrame** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例使第一个形状的文本框能自动调整大小，以适应其中所包含的文字。*/ Application.Worksheets.Item(1).Shapes.Item(1).TextFrame.AutoSize = true` |

#### **TextFrame.HorizontalAnchor**

**语法**

**express.HorizontalAnchor**

*express*   一个代表 **TextFrame** 对象的变量。

#### **TextFrame.MarginBottom**

以磅为单位返回或设置从文本框底端到包含文本的形状中内接矩形底端的距离。可读/写。**Single** 类型。

**语法**

**express.MarginBottom**

*express*   一个代表 **TextFrame** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在 myDocument 中添加一个矩形，向矩形中添加文本，然后设置文本框架的边距。*/ function test() {   let myDocument = Application.Worksheets.Item(1)   let rng = myDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame   rng.Characters().Text = "Here is some test text"   rng.MarginBottom = 0   rng.MarginLeft = 100   rng.MarginRight = 0   rng.MarginTop = 20 }` |

#### **TextFrame.MarginLeft**

以磅为单位返回或设置从文本框左边界到包含文本的形状中内接矩形左边界的距离。可读写。**Single** 类型。

**语法**

**express.MarginLeft**

*express*   一个代表 **TextFrame** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在 myDocument 中添加一个矩形，向矩形中添加文本，然后设置文本框架的边距。*/ function test() {   let myDocument = Application.Worksheets.Item(1)   let rng = myDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame   rng.Characters().Text = "Here is some test text"   rng.MarginBottom = 0   rng.MarginLeft = 100   rng.MarginRight = 0   rng.MarginTop = 20 }` |

#### **TextFrame.MarginRight**

以磅为单位返回或设置从文本框右边界到包含文本的形状中内接矩形右边界的距离。可读写。**Single** 类型。

**语法**

**express.MarginRight**

*express*   一个代表 **TextFrame** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在 myDocument 中添加一个矩形，向矩形中添加文本，然后设置文本框架的边距。*/ function test() {   let myDocument = Application.Worksheets.Item(1)   let rng = myDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame   rng.Characters().Text = "Here is some test text"   rng.MarginBottom = 0   rng.MarginLeft = 100   rng.MarginRight = 0   rng.MarginTop = 20 }` |

#### **TextFrame.MarginTop**

以磅为单位返回或设置从文本框架顶端到包含文本的形状中内接矩形顶端的距离。可读写。**Single** 类型。

**语法**

**express.MarginTop**

*express*   一个代表 **TextFrame** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在 myDocument 中添加一个矩形，向矩形中添加文本，然后设置文本框架的边距。*/ function test() {   let myDocument = Application.Worksheets.Item(1)   let rng = myDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame   rng.Characters().Text = "Here is some test text"   rng.MarginBottom = 0   rng.MarginLeft = 100   rng.MarginRight = 0   rng.MarginTop = 20 }` |

#### **TextFrame.Orientation**

返回或设置一个 **Long** 值，它代表文本框的方向。

**语法**

**express.Orientation**

*express*   一个代表 **TextFrame** 对象的变量。

**说明**

此属性值可设为 -90 到 90 度之间的整数值或以下 **MsoTextOrientation** 常量之一。

#### **TextFrame.Ruler**

**语法**

**express.Ruler**

*express*   一个代表 **TextFrame** 对象的变量。

#### **TextFrame.TextRange**

**语法**

**express.TextRange**

*express*   一个代表 **TextFrame** 对象的变量。

#### **TextFrame.VerticalAnchor**

**语法**

**express.VerticalAnchor**

*express*   一个代表 **TextFrame** 对象的变量。

#### **TextFrame.WordWrap**

**语法**

**express.WordWrap**

*express*   一个代表 **TextFrame** 对象的变量。

适用环境：web

适用平台：windows/linux