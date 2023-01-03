**HeaderFooter**



代表幻灯片或母版上的页眉、页脚、日期和时间、幻灯片编号或页码。幻灯片或母版的所有 **HeaderFooter** 对象都包含在 **HeadersFooters** 对象中。

**说明**

使用下表中列出的属性之一返回 **HeaderFooter** 对象。

| 使用此属性  | 返回                                                         |
| ----------- | ------------------------------------------------------------ |
| DateAndTime | 代表幻灯片的日期和时间的 HeaderFooter 对象。                 |
| Footer      | 代表幻灯片页脚的 HeaderFooter 对象。                         |
| Header      | 代表幻灯片页眉的 HeaderFooter 对象。仅对备注页和讲义有效，不用于幻灯片。 |
| SlideNumber | 代表幻灯片编号（在幻灯片中）或页码（在备注页或讲义中）的 HeaderFooter 对象。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个 **Application** 对象，该对象表示指定对象的创建者。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Format**      | 返回或设置自动更新的日期和时间的格式。仅适用于 **HeaderFooter** 对象，该对象代表时间和日期（由 **DateAndTime** 属性从 **HeadersFooters** 集合中返回）。可读/写 **PpDateTimeFormat** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Text**        | 返回或设置一个 **String** 类型值，该值代表指定对象中包含的文本。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **UseFormat**   | 决定日期和时间对象是否包含自动更新的信息。可读/写 **MsoTriState** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Visible**     | 返回或设置指定对象的可见性或应用于指定对象的格式。可读/写。**MsoTriState** 类型。 |

**成员属性**

#### **HeaderFooter.Application**

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **HeaderFooter** 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function AddAndSave(pptPre) {     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") }` |

以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes     for (let i = 1; i <= shpOle.Count; i++) {         if (shpOle.Item(i).Type == msoLinkedOLEObject) {             alert(shpOle.Item(i).OLEFormat.Application.Name)         }     } }` |

#### **HeaderFooter.Format**

返回或设置自动更新的日期和时间的格式。仅适用于 **HeaderFooter** 对象，该对象代表时间和日期（由 **DateAndTime** 属性从 **HeadersFooters** 集合中返回）。可读/写 **PpDateTimeFormat** 类型。

**语法**

**express.Format**

*express*   一个代表 **HeaderFooter** 对象的变量。

**说明**

通过将 **UseFormat** 属性设置为 **True**，确保将日期和时间设置为自动更新（不作为固定文本显示）

| PpDateTimeFormat 可以是下列 PpDateTimeFormat 常量之一。 |
| ------------------------------------------------------- |
| ppDateTimeddddMMMMddyyyy                                |
| ppDateTimedMMMMyyyy                                     |
| ppDateTimedMMMyy                                        |
| ppDateTimeFormatMixed                                   |
| ppDateTimeHmm                                           |
| ppDateTimehmmAMPM                                       |
| ppDateTimeHmmss                                         |
| ppDateTimehmmssAMPM                                     |
| ppDateTimeMdyy                                          |
| ppDateTimeMMddyyHmm                                     |
| ppDateTimeMMddyyhmmAMPM                                 |
| ppDateTimeMMMMdyyyy                                     |
| ppDateTimeMMMMyy                                        |
| ppDateTimeMMyy                                          |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将当前演示文稿的幻灯片母版中的日期和时间设为自动更新，并将日期和时间格式设为显示小时、分钟和秒。*/ function test(){ 　　　　let da = Application.ActivePresentation.SlideMaster.HeadersFooters.DateAndTime 　　　　da.UseFormat = true 　　　　da.Format = ppDateTimeHmmss }` |

#### **HeaderFooter.Parent**

返回指定对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **HeaderFooter** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/ function test() {     let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes     let addShp = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame     addShp.TextRange.Text = "Test text"     addShp.Parent.Rotation = 45 }` |

#### **HeaderFooter.Text**

返回或设置一个 **String** 类型值，该值代表指定对象中包含的文本。可读/写。

**语法**

**express.Text**

*express*   一个代表 **HeaderFooter** 对象的变量。

#### **HeaderFooter.UseFormat**

决定日期和时间对象是否包含自动更新的信息。可读/写 **MsoTriState** 类型。

**语法**

**express.UseFormat**

*express*   一个代表 **HeaderFooter** 对象的变量。

**说明**

此属性仅适用于 **HeaderFooter** 对象，该对象代表日期和时间（由 **DateAndTime** 属性返回）。如果希望使用 **Format** 属性来设置或返回日期合时间的格式，请将日期和时间的 **HeaderFooter** 对象的 **UseFormat** 属性设为 **True**。如果希望设置或返回固定日期和时间的文本字符串，请将 **UseFormat** 属性设为 **msoFalse**。

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| --------------------------------------------- |
| msoCTrue                                      |
| msoFalse 日期和时间对象是固定字符串。         |
| msoTriStateMixed                              |
| msoTriStateToggle                             |
| msoTrue 日期和时间对象包含自动更新的信息。    |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将当前演示文稿的幻灯片母版中的日期和时间设为自动更新，并将日期和时间格式设为显示小时、分钟和秒。*/ function test(){ 　　　　let da = Application.ActivePresentation.SlideMaster.HeadersFooters.DateAndTime 　　　　da.UseFormat = msoTrue 　　　　da.Format = ppDateTimeHmmss }` |

#### **HeaderFooter.Visible**

返回或设置指定对象的可见性或应用于指定对象的格式。可读/写。**MsoTriState** 类型。

**语法**

**express.Visible**

*express*   一个代表 **HeaderFooter** 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| --------------------------------------------- |
| msoCTrue                                      |
| msoFalse                                      |
| msoTriStateMixed                              |
| msoTriStateToggle                             |
| msoTrue：指定对象或对象格式是可见的。         |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例为当前演示文稿中第一张幻灯片的第三个形状设置阴影的水平偏移量和垂直偏移量，使其向右偏移 5 磅并向上偏移 3 磅。如果该形状原来没有阴影，本例将会为其添上阴影。*/ function test(){ 　　　　let sh= ActivePresentation.Slides.Item(1).Shapes.Item(3).Shadow 　　　　sh.Visible = msoTrue 　　　　sh.OffsetX = 5 　　　　sh.OffsetY = -3 }` |

适用环境：web

适用平台：windows/linux