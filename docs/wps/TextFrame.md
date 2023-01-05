#### **TextFrame**



代表 **Shape** 对象中的文本框。**TextFrame** 对象包含文本框架中的文本以及用于控制文本框架的边距和方向的属性。

**说明**

使用 **TextFrame** 属性可返回形状的 **TextFrame** 对象。**TextRange** 属性可返回一个 [Range ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象，该对象代表指定文本框架中的文本范围。以下示例在活动文档中第一个形状的文本框架中添加文本。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActiveDocument.Shapes.Item(1).TextFrame.TextRange.Text = "My Text"` |

| ![img](https://qn.cache.wpscdn.cn/gif/close.gif)注释         |
| ------------------------------------------------------------ |
| 有些形状不支持附加文本（如线条、任意多边形、图片和 OLE 对象）。如果试图返回或设置用于控制这些对象的文本框架中文本的属性，那么将发生错误。 |

使用 **HasText** 属性可确定文本框架中是否包含文本，如以下示例所示。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `for(let i = 1; i <= Application.ActiveDocument.Shapes.Count; i++){     if(Application.ActiveDocument.Shapes.Item(i).TextFrame.HasText){         alert(Application.ActiveDocument.Shapes.Item(i).TextFrame.TextRange.Text)     } }` |

文本框架可以链接在一起，以使一个形状的文本框架中的文本排入另一个形状的文本框架。使用 **Next** 和 **Previous** 属性可链接文本框架。以下示例创建一个文本框（一个带有文本框架的矩形），并在其中添加文本。然后创建另一个文本框，并将两个文本框架链接在一起，以使第一个文本框架中的文本排入第二个文本框架。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `let myTB1 = Application.ActiveDocument.Shapes.AddTextbox(msoTextOrientationHorizontal, 72, 72, 72, 36)     myTB1.TextFrame.TextRange.Text = "This is some text. This is some more text."  let myTB2 = Application.ActiveDocument.Shapes.AddTextbox(msoTextOrientationHorizontal, 144, 144, 72, 36)     myTB1.TextFrame.Next = myTB2.TextFrame` |

可以使用 **ContainingRange** 属性返回 **Range** 对象，该对象代表在链接的文本框架之间排列的整个文字部分。以下示例检查文本框 3 中的文本和链接到文本框 3 的其他所有文本的拼写。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `let myStory = Application.ActiveDocument.Shapes.Item("TextBox 3").TextFrame.ContainingRange myStory.CheckSpelling()` |

**方法**

|                                                              | 名称                 | 说明                                                         |
| ------------------------------------------------------------ | -------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **BreakForwardLink** | 在链接存在的情况下，断开指定文本框的前向链接。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **DeleteText**       | 删除文本框中的文本以及该文本的所有关联属性，包括字体属性。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ValidLinkTarget**  | 确定一个形状的文本框是否可以链接到另一个形状的文本框。返回Boolean值 |

**属性**

|                                                              | 名称                 | 说明                                                         |
| ------------------------------------------------------------ | -------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**      | 返回一个代表 WPS 应用程序的 [Application ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoSize**         | 返回或设置一个 **Long** 类型的值，该值代表是否自动调整文本框的大小。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Column**           | 返回代表指定文本框架的列的 [Column ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Column/Column%20.htm#jsObject_Column)对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ContainingRange**  | 返回一个 [Range ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象，该对象代表一系列带有链接文本框架（包括指定文本框架）的图形的整个文字部分。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**          | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HasText**          | 如果指定的图形有相关的文字，则该属性值为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HorizontalAnchor** | 返回或设置文本框架中文本的水平对齐方式。可读/写 **MsoHorizontalAnchor**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MarginBottom**     | 以磅为单位返回或设置从文本框底端到包含文本的形状中内接矩形底端的距离。可读写。**Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MarginLeft**       | 以磅为单位返回或设置从文本框左边界到包含文本的形状中内接矩形左边界的距离。可读写。**Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MarginRight**      | 以磅为单位返回或设置从文本框右边界到包含文本的形状中内接矩形右边界的距离。可读写。**Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MarginTop**        | 以磅为单位返回或设置从文本框架顶端到包含文本的形状中内接矩形顶端的距离。可读写。**Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Next**             | 返回一个 **TextFrame** 对象，该对象代表形状集合中的下一个文本框架。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NoTextRotation**   | 如果在旋转形状时不应旋转文本框架中的文本，则该属性值为 True。可读/写 MsoTriState。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Orientation**      | 返回或设置图文框内的文字方向。可读写 **MsoTextOrientation** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Overflowing**      | 如果指定文本框架中的文本溢出该框架，则该属性值为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**           | 返回一个 **Shape** 对象，该对象代表文本框架的父形状。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PathFormat**       | 返回或设置指定文本框架的路径类型。可读/写 **MsoPathType**。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Previous**         | 返回一个 **TextFrame** 对象，该对象代表形状集合中的上一个文本框架。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextRange**        | 返回一个 [Range ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象，该对象代表指定文本框架中的文本。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ThreeD**           | 返回**ThreeDFormat**对象，该对象包含指定文本框架的三维效果格式属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **VerticalAnchor**   | 返回或设置一个 **MsoVerticalAnchor** 常量，该常量代表形状内文本的垂直对齐方式。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WarpFormat**       | 返回或设置指定文本框架的环绕格式（文本如何环绕）。可读/写 **MsoWarpFormat**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WordWrap**         | 如果 WPS 在指定文本框架的西文单词中间断字换行，则该属性值为 **True**。**Long** 类型，可读写。 |

**成员方法**

#### **TextFrame.BreakForwardLink**

在链接存在的情况下，断开指定文本框的前向链接。

**语法**

**express.BreakForwardLink()**

*express*   一个代表 **TextFrame** 对象的变量。

**说明**

如果将此方法应用于图形，且该图形位于一串与图文框相链接的图形的中间位置，则会使链接断开，而形成两组链接的图形。但所有文本都将保留在第一组链接的图形中。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：新建一个文档，在文档中添加三个前后链接的文本框，然后断开第二个文本框后面的链接。*/ function test() {     Documents.Add()     let shapeTextBox1 = Application.ActiveDocument.Shapes.AddTextbox(msoTextOrientationHorizontal, InchesToPoints(1.5), InchesToPoints(0.5), InchesToPoints(1), InchesToPoints(0.5))     shapeTextBox1.TextFrame.TextRange.Text = "This is some text. This is some more text. This is even more text."      let shapeTextBox2 = Application.ActiveDocument.Shapes.AddTextbox(msoTextOrientationHorizontal, InchesToPoints(1.5), InchesToPoints(1.5), InchesToPoints(1), InchesToPoints(0.5))     let shapeTextBox3 = Application.ActiveDocument.Shapes.AddTextbox(msoTextOrientationHorizontal, InchesToPoints(1.5), InchesToPoints(2.5), InchesToPoints(1), InchesToPoints(0.5))      shapeTextBox1.TextFrame.Next = shapeTextBox2.TextFrame     shapeTextBox2.TextFrame.Next = shapeTextBox3.TextFrame      alert("TextBoxes 1, 2, and 3 are linked.")      shapeTextBox2.TextFrame.BreakForwardLink() }` |

#### **TextFrame.DeleteText**

删除文本框中的文本以及该文本的所有关联属性，包括字体属性。

**语法**

**express.DeleteText()**

*express*   一个代表 **TextFrame** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例从活动文档的第一个形状中删除文本（如果该形状包含文本）。*/ function DeleteText_Example(){     Application.ActiveDocument.Shapes.Item(1).TextFrame.DeleteText() }` |

#### **TextFrame.ValidLinkTarget**

确定一个形状的文本框是否可以链接到另一个形状的文本框。返回Boolean值

**语法**

**express.ValidLinkTarget(TargetTextFrame)**

*express*   一个代表 **TextFrame** 对象的变量。

**参数**

| **名称**          | **必选/可选** | **数据类型**  | **说明**                                    |
| ----------------- | ------------- | ------------- | ------------------------------------------- |
| *TargetTextFrame* | 必选          | **TextFrame** | expression 返回的文本框要链接的目标文本框。 |

**说明**

如果 *TargetTextFrame* 是有效目标，则该方法返回 **True**；如果 *TargetTextFrame* 已包含文字或已链接，或者该形状不支持附加文字，则返回 **False**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：检查活动文档第一和第二个图形的文本框，以判断其是否能相互链接。如果能，则链接两个文本框。*/ let textFrame1 = Application.ActiveDocument.Shapes.Item(1).TextFrame let textFrame2 = Application.ActiveDocument.Shapes.Item(2).TextFrame  if(textFrame1.ValidLinkTarget(textFrame2) == true){     textFrame1.Next = textFrame2 }` |

**成员属性**

#### **TextFrame.Application**

返回一个代表 WPS 应用程序的 [Application ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。

**语法**

**express.Application**

*express*   一个代表 **TextFrame** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **TextFrame.AutoSize**

返回或设置一个 **Long** 类型的值，该值代表是否自动调整文本框的大小。可读写。

**语法**

**express.AutoSize**

*express*   一个代表 **TextFrame** 对象的变量。

#### **TextFrame.Column**

返回代表指定文本框架的列的 [Column ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Column/Column%20.htm#jsObject_Column)对象。只读。

**语法**

**express.Column**

*express*   一个代表 **TextFrame** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function Column_Example(){     Application.ActiveDocument.Shapes.Item(1).TextFrame.Column.Number = 2 }` |

#### **TextFrame.ContainingRange**

返回一个 [Range ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象，该对象代表一系列带有链接文本框架（包括指定文本框架）的图形的整个文字部分。只读。

**语法**

**express.ContainingRange**

*express*   一个代表 **TextFrame** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对 TextBox 1 及所有链接至 TextBox 1 的文本框架的文本进行拼写检查。*/ let rngStory = Application.ActiveDocument.Shapes.Item("TextBox 1").TextFrame.ContainingRange rngStory.CheckSpelling()` |

#### **TextFrame.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **TextFrame** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **TextFrame.HasText**

如果指定的图形有相关的文字，则该属性值为 **True**。**Boolean** 类型，只读。

**语法**

**express.HasText**

*express*   一个代表 **TextFrame** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果活动文档中的第二个图形包含文字，并且文字超出了图文框时，本示例显示一条消息。*/ let docActive = Application.ActiveDocument  if(docActive.Shapes.Item(2).TextFrame.HasText == true){     if(docActive.Shapes.Item(2).TextFrame.Overflowing == true){         MsgBox("Text overflows the frame.")     } }` |

#### **TextFrame.HorizontalAnchor**

返回或设置文本框架中文本的水平对齐方式。可读/写 **MsoHorizontalAnchor**。

**语法**

**express.HorizontalAnchor**

*express*   一个代表 **TextFrame** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例演示如何将活动文档中第一个形状的对齐方式设置为靠上居中。*/ function HorizontalAnchor_Example(){     Application.ActiveDocument.Shapes.Item(1).TextFrame.HorizontalAnchor = msoAnchorCenter     Application.ActiveDocument.Shapes.Item(1).TextFrame.VerticalAnchor = msoAnchorTop }` |

#### **TextFrame.MarginBottom**

以磅为单位返回或设置从文本框底端到包含文本的形状中内接矩形底端的距离。可读写。**Single** 类型。

**语法**

**express.MarginBottom**

*express*   一个代表 **TextFrame** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在 myDocument 中添加一个矩形，向矩形中添加文本，然后设置文本框架的边距。*/ function test() {     let myShape = Application.ActiveDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame      myShape.TextRange.Text = "Here is some test text"     myShape.MarginBottom = 100     myShape.MarginLeft = 100     myShape.MarginRight = 50     myShape.MarginTop = 20 } ` |

#### **TextFrame.MarginLeft**

以磅为单位返回或设置从文本框左边界到包含文本的形状中内接矩形左边界的距离。可读写。**Single** 类型。

**语法**

**express.MarginLeft**

*express*   一个代表 **TextFrame** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在 myDocument 中添加一个矩形，向矩形中添加文本，然后设置文本框架的边距。*/ let myShape = Application.ActiveDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame  myShape.TextRange.Text = "Here is some test text" myShape.MarginBottom = 100 myShape.MarginLeft = 100 myShape.MarginRight = 50 myShape.MarginTop = 20` |

#### **TextFrame.MarginRight**

以磅为单位返回或设置从文本框右边界到包含文本的形状中内接矩形右边界的距离。可读写。**Single** 类型。

**语法**

**express.MarginRight**

*express*   一个代表 **TextFrame** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在 myDocument 中添加一个矩形，向矩形中添加文本，然后设置文本框架的边距。*/ let myShape = Application.ActiveDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame  myShape.TextRange.Text = "Here is some test text" myShape.MarginBottom = 100 myShape.MarginLeft = 100 myShape.MarginRight = 50 myShape.MarginTop = 20` |

#### **TextFrame.MarginTop**

以磅为单位返回或设置从文本框架顶端到包含文本的形状中内接矩形顶端的距离。可读写。**Single** 类型。

**语法**

**express.MarginTop**

*express*   一个代表 **TextFrame** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在 myDocument 中添加一个矩形，向矩形中添加文本，然后设置文本框架的边距。*/ let myShape = Application.ActiveDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame  myShape.TextRange.Text = "Here is some test text" myShape.MarginBottom = 100 myShape.MarginLeft = 100 myShape.MarginRight = 50 myShape.MarginTop = 20` |

#### **TextFrame.Next**

返回一个 **TextFrame** 对象，该对象代表形状集合中的下一个文本框架。只读。

**语法**

**express.Next**

*express*   一个代表 **TextFrame** 对象的变量。

#### **TextFrame.NoTextRotation**

如果在旋转形状时不应旋转文本框架中的文本，则该属性值为 True。可读/写 MsoTriState。

**语法**

**express.NoTextRotation**

*express*   一个代表 **TextFrame** 对象的变量。

**说明**

仅可使用 **MsoTriState** 常量 **msoFalse** 和 **msoTrue** 来设置该属性。

#### **TextFrame.Orientation**

返回或设置图文框内的文字方向。可读写 **MsoTextOrientation** 类型。

**语法**

**express.Orientation**

*express*   一个代表 **TextFrame** 对象的变量。

**说明**

根据您所选择或安装的语言支持（例如，美国英语），某些 **MsoTextOrientation** 常量可能不可用。

您可以设置文本框架或者恰好位于文本框架内的范围或所选内容的方向。有关文本框架和文本框之间的区别的信息，请参阅 **TextFrame** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例新建一篇文档，在文档中插入文字，接着使用这些文字创建一个文本框，然后设置该文本框架的方向，使其中的文字向上倾斜。*/ function test() {     let myDoc = Application.Documents.Add()     Application.Selection.TypeText("This is some text.")     myDoc.Content.Select()     Application.Selection.CreateTextbox()     myDoc.Shapes.Item(1).TextFrame.Orientation = msoTextOrientationUpward } ` |

#### **TextFrame.Overflowing**

如果指定文本框架中的文本溢出该框架，则该属性值为 **True**。**Boolean** 类型，只读。

**语法**

**express.Overflowing**

*express*   一个代表 **TextFrame** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例检查 MyTextBox 中的文本是否超出其文本框架。如果超出，本示例将添加另一个文本框架并链接这两个文本框，以使文本排入第二个文本框架。*/ let myTBox = Application.ActiveDocument.Shapes.Item("MyTextBox")  if(myTBox.TextFrame.Overflowing == true){     let nextTBox = Application.ActiveDocument.Shapes.AddTextbox(msoTextOrientationHorizontal, 72, 72, 100, 200)     myTBox.TextFrame.Next = nextTBox.TextFrame } ` |

#### **TextFrame.Parent**

返回一个 **Shape** 对象，该对象代表文本框架的父形状。

**语法**

**express.Parent**

*express*   一个代表 **TextFrame** 对象的变量。

#### **TextFrame.PathFormat**

返回或设置指定文本框架的路径类型。可读/写 **MsoPathType**。

**语法**

**express.PathFormat**

*express*   一个代表 **TextFrame** 对象的变量。

**说明**

**PathFormat** 属性的值可以是下列 **MsoPathType** 常量之一：

- msoPathType1
- msoPathType2
- msoPathType3
- msoPathType4
- msoPathTypeMixed
- msoPathTypeNone

无法设置值 **msoPathTypeMixed**。如果设置值 **msoPathTypeNone**，则会删除所有现有路径。

#### **TextFrame.Previous**

返回一个 **TextFrame** 对象，该对象代表形状集合中的上一个文本框架。只读。

**语法**

**express.Previous**

*express*   一个代表 **TextFrame** 对象的变量。

#### **TextFrame.TextRange**

返回一个 [Range ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象，该对象代表指定文本框架中的文本。

**语法**

**express.TextRange**

*express*   一个代表 **TextFrame** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动文档中添加一个文本框，然后在文本框中添加文字。*/ let myTBox = Application.ActiveDocument.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 300, 200) myTBox.TextFrame.TextRange.Text = "Test Box"  /*本示例在活动文档的文本框 1 中添加文字。*/ Application.ActiveDocument.Shapes.Item("TextBox 1").TextFrame.TextRange.InsertAfter("New Text")  /*本示例返回活动文档的文本框 1 中的文字，并显示在消息框中。*/ alert(ActiveDocument.Shapes.Item("TextBox 1").TextFrame.TextRange.Text)` |

#### **TextFrame.ThreeD**

返回**ThreeDFormat**对象，该对象包含指定文本框架的三维效果格式属性。只读。

**语法**

**express.ThreeD**

*express*   一个代表 **TextFrame** 对象的变量。

#### **TextFrame.VerticalAnchor**

返回或设置一个 **MsoVerticalAnchor** 常量，该常量代表形状内文本的垂直对齐方式。可读写。

**语法**

**express.VerticalAnchor**

*express*   一个代表 **TextFrame** 对象的变量。

#### **TextFrame.WarpFormat**

返回或设置指定文本框架的环绕格式（文本如何环绕）。可读/写 **MsoWarpFormat**。

**语法**

**express.WarpFormat**

*express*   一个代表 **TextFrame** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例演示如何设置活动文档中第一个形状的环绕格式。*/ Application.ActiveDocument.Shapes.Item(1).TextFrame.WarpFormat = msoWarpFormat15` |

#### **TextFrame.WordWrap**

如果 WPS 在指定文本框架的西文单词中间断字换行，则该属性值为 **True**。**Long** 类型，可读写。

**语法**

**express.WordWrap**

*express*   一个代表 **TextFrame** 对象的变量。

**说明**

如果该属性对指定文本框架中的某些指定文本设置为 **True** 而对其他文本设置为 false，则该属性返回 **wdUndefined**。

| ![img](https://qn.cache.wpscdn.cn/gif/close.gif)注释         |
| ------------------------------------------------------------ |
| 根据您所选择或安装的语言支持（例如，美国英语），该属性可能不可用。 |

适用环境：web

适用平台：windows/linux