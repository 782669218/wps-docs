**Hyperlinks**



代表工作表或区域的超链接的集合。

**说明**

每个超链接都由一个 **Hyperlink** 对象代表。

使用 **Hyperlinks** 属性可返回 **Hyperlinks** 集合。下例检查工作表一上的超链接，看是否有包含“Microsoft”一词的链接。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     for(let h = 1; h <= Application.Worksheets.Item(1).Hyperlinks.Count; h++) {         if(Application.Worksheets.Item(1).Hyperlinks.Item(h).Name.indexOf("Microsoft") != -1) {  Application.Worksheets.Item(1).Hyperlinks.Item(h).Follow()         }     } }` |

使用 [**Add** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Hyperlinks/Hyperlinks%20.htm#Hyperlinks.Add)方法可创建一个超链接并将它添加到 **Hyperlinks** 集合。下例为单元格 E5 创建一个新的超链接。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let hyperlinks2 = Application.Worksheets.Item(1)     hyperlinks2.Hyperlinks.Add(sheet2.Range("E5"), "http://example.microsoft.com") } ` |

**方法**

|                                                              | 名称       | 说明                                                         |
| ------------------------------------------------------------ | ---------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**    | 向指定的区域或形状添加超链接。返回一个 **Hyperlink** 对象，它代表新的超链接。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete** | 删除对象。                                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**   | 从集合中返回一个对象。                                       |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Long** 值，它代表集合中对象的数量。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |

**成员方法**

#### **Hyperlinks.Add**

向指定的区域或形状添加超链接。返回一个 **Hyperlink** 对象，它代表新的超链接。

**语法**

**express.Add(Anchor, Address, SubAddress, ScreenTip, TextToDisplay)**

*express*   一个代表 **Hyperlinks** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型** | **说明**                                     |
| --------------- | ------------- | ------------ | -------------------------------------------- |
| *Anchor*        | 必选          | **Object**   | 超链接的位置。可为 Range 或 Shape 对象。     |
| *Address*       | 必选          | **String**   | 超链接的地址。                               |
| *SubAddress*    | 可选          | **Variant**  | 超链接的子地址。                             |
| *ScreenTip*     | 可选          | **Variant**  | 当鼠标指针停留在超链接上时所显示的屏幕提示。 |
| *TextToDisplay* | 可选          | **Variant**  | 要显示的超链接的文本。                       |

**说明**

指定 **TextToDisplay** 参数时，文本必须是字符串。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     /* 向单元格 A5 添加超链接。*/     let sheet2 = Application.Worksheets.Item(1)     sheet2.Hyperlinks.Add(sheet2.Range("a5"), "http://example.microsoft.com", null, "Microsoft Web Site", "Microsoft")      /* 向单元格 A5 添加一个电子邮件超链接。*/     let sheet2 = Application.Worksheets.Item(1)     sheet2.Hyperlinks.Add(sheet2.Range("a5"), "mailto:someone@example.com?subject=hello", null, "Write us today", "Support") }` |

#### **Hyperlinks.Delete**

删除对象。

**语法**

**express.Delete()**

*express*   一个代表 **Hyperlinks** 对象的变量。

#### **Hyperlinks.Item**

从集合中返回一个对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **Hyperlinks** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**             |
| -------- | ------------- | ------------ | -------------------- |
| *Index*  | 必选          | **Variant**  | 对象的名称或索引号。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例激活 E5 单元格的第一个超链接。*/ Application.Worksheets.Item(1).Range("E5").Hyperlinks.Item(1).Follow()` |

**成员属性**

#### **Hyperlinks.Application**

如果不使用对象识别符，则该属性返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **Hyperlinks** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() {     let myObject = Application.ActiveWorkbook     if(myObject.Application.Value == "ET") {         alert("This is an ET Application object.")     }     else {         alert("This is not an ET Application object.")     } }` |

#### **Hyperlinks.Count**

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

*express*   一个代表 **Hyperlinks** 对象的变量。

#### **Hyperlinks.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Hyperlinks** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **Hyperlinks.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **Hyperlinks** 对象的变量。

适用环境：web

适用平台：windows/linux