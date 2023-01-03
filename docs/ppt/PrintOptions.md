**PrintOptions**



包含演示文稿的打印选项。

注释：指定 **PrintOut**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Document/Document%20.htm#Document.PrintOut)方法的可选参数 **From**、**To**、**Copies** 和 **Collate** 将设置 **PrintOptions** 对象的相应属性。

**说明**

使用 **PrintOptions** 属性返回 **PrintOptions** 对象。以下示例以非逐份方式打印活动演示文稿所有幻灯片（无论可见或隐藏）的两个彩色副本。该示例还调整每个幻灯片的大小以适应打印页，并给每个幻灯片加细边框。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){     let active = Application.ActivePresentation         let print = active.PrintOptions             print.NumberOfCopies = 2             print.Collate = false             print.PrintColorType = ppPrintColor             print.PrintHiddenSlides = true             print.FitToPage = true             print.FrameSlides = true             print.OutputType = ppPrintOutputSlides         active.PrintOut() } ` |

使用 **RangeType**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/PrintOptions/PrintOptions%20.htm#PrintOptions.RangeType)属性指定打印整个演示文稿或仅打印其中的指定部分。如果想要只打印特定幻灯片，请将 **RangeType**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/PrintOptions/PrintOptions%20.htm#PrintOptions.RangeType)属性设为 **ppPrintSlideRange**，并使用 **Ranges**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/PrintOptions/PrintOptions%20.htm#PrintOptions.Ranges)属性指定要打印的页。以下示例打印活动演示文稿中的第一、四、五、六张幻灯片。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){     let active = Application.ActivePresentation         let print = active.PrintOptions             print.RangeType = ppPrintSlideRange             let ranges = print.Ranges                 ranges.Add(1, 1)                 ranges.Add(4, 6)         active.PrintOut() } ` |

**属性**

|                                                              | 名称                     | 说明                                                         |
| ------------------------------------------------------------ | ------------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ActivePrinter**        | 返回当前打印机的名称。只读。**String** 类型。                |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**          | 返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示指定对象的创建者。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Collate**              | 决定是否在打印指定演示文稿的完整副本之后才打印其下一个副本的第一页。可读/写 **MsoTriState** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FitToPage**            | 决定是否调整幻灯片大小以适应打印页面。**MsoTriState** 类型，可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FrameSlides**          | 决定是否在打印的幻灯片的边框周围放置窄框架。可读/写 **MsoTriState** 类型。适用于打印的幻灯片、讲义和备注页。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HandoutOrder**         | 返回或设置在打印的讲义上显示幻灯片的页面版式顺序，其中可在一个页面上显示多张幻灯片。可读/写 **PpPrintHandoutOrder** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NumberOfCopies**       | 返回或设置要打印的演示文稿份数。默认值为 1。可读/写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OutputType**           | 返回或设置一个值，该值指示要打印演示文稿的哪些组件（幻灯片、讲义、备注页或大纲）。可读/写 **PpPrintOutputType** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**               | 返回指定对象的父对象。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PrintColorType**       | 返回或设置指定文档的打印方式：黑白方式、纯黑色和纯白色（亦称为高对比度）方式或彩色方式。默认值由打印机设置。可读/写 **PpPrintColorType** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PrintComments**        | 设置或返回是否打印批注。可读/写 **MsoTriState** 类型。       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PrintFontsAsGraphics** | 决定是否将 TrueType 字体作为图形打印。可读/写 **MsoTriState** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PrintHiddenSlides**    | 决定是否打印指定演示文稿中的隐藏幻灯片。可读/写 **MsoTriState** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PrintInBackground**    | 决定是否在后台打印指定的演示文稿。可读/写 **MsoTriState** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Ranges**               | 返回一个 **PrintRanges** 对象，该对象代表演示文稿中要打印的幻灯片范围。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RangeType**            | 返回或设置演示文稿打印范围的类型。可读/写 **PpPrintRangeType** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SlideShowName**        | 返回或设置自定义幻灯片放映的名称，该幻灯片放映将运行以响应幻灯片放映期间在形状上的鼠标动作（**ActionSetting** 对象）；返回或设置自定义幻灯片放映的名称以便打印（**PrintOptions** 对象）；返回或设置发布为 Web 演示文稿的自定义幻灯片放映的名称（**PublishObject** 对象）。**String** 类型，可读/写。 |

**成员属性**

#### **PrintOptions.ActivePrinter**

返回当前打印机的名称。只读。**String** 类型。

**语法**

**express.ActivePrinter**

*express*   一个代表 **PrintOptions** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示当前打印机的名称。*/ alert("The name of the active printer is " + Application.ActivePrinter)` |

#### **PrintOptions.Application**

返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **PrintOptions** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。*/ function test(pptPres) {     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") } /*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。*/ function test() {   let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes   for(let i = 1; i <= myShapes.Count; i++){       if(myShapes.Item(i).Type == msoLinkedOLEObject){           alert(myShapes.Item(i).OLEFormat.Application.Name)       }   } }` |

#### **PrintOptions.Collate**

决定是否在打印指定演示文稿的完整副本之后才打印其下一个副本的第一页。可读/写 **MsoTriState** 类型。

**语法**

**express.Collate**

*express*   一个代表 **PrintOptions** 对象的变量。

**说明**

指定 **PrintOut** 方法的 **Collate** 参数值将设置该属性的值。

| MsoTriState 可以是下列 MsoTriState 常量之一。                |
| ------------------------------------------------------------ |
| **msoCTrue**                                                 |
| **msoFalse**                                                 |
| **msoTriStateMixed**                                         |
| **msoTriStateToggle**                                        |
| **msoTrue** 默认值。在打印指定演示文稿的完整副本之后才打印其下一个副本的第一页。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例将活动演示文稿逐份打印三份。 function CheckOutPresentation(strPresentation) {     let print = Application.ActivePresentation.PrintOptions     print.NumberOfCopies = 3     print.Collate = msoTrue     print.Parent.PrintOut() }` |

#### **PrintOptions.FitToPage**

决定是否调整幻灯片大小以适应打印页面。**MsoTriState** 类型，可读/写。

**语法**

**express.FitToPage**

*express*   一个代表 **PrintOptions** 对象的变量。

**说明**

| **MsoTriState** 可以是下列 **MsoTriState** 常量之一。        |
| ------------------------------------------------------------ |
| **msoCTrue**                                                 |
| **msoFalse** 默认值。幻灯片将使用“页面设置”对话框中指定的尺寸，不论该尺寸是否与打印页面相适应。 |
| **msoTriStateMixed**                                         |
| **msoTriStateToggle**                                        |
| **msoTrue** 调整指定幻灯片的大小以适应打印页面，而不考虑“页面设置”对话框（“文件”菜单）中“高度”框和“宽度”框中的值。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将打印当前演示文稿并调整每张幻灯片的大小以适应打印页面。*/ function test() {   let active = Application.ActivePresentation       active.PrintOptions.FitToPage = msoTrue       active.PrintOut() }` |

#### **PrintOptions.FrameSlides**

决定是否在打印的幻灯片的边框周围放置窄框架。可读/写 **MsoTriState** 类型。适用于打印的幻灯片、讲义和备注页。

**语法**

**express.FrameSlides**

*express*   一个代表 **PrintOptions** 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。    |
| ------------------------------------------------ |
| **msoCTrue**                                     |
| **msoFalse** 默认值。                            |
| **msoTriStateMixed**                             |
| **msoTriStateToggle**                            |
| **msoTrue** 在打印的幻灯片的边框周围放置窄框架。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例打印活动演示文稿，其中每张幻灯片周围都有一个框架。 function test() {     let active = Application.ActivePresentation     active.PrintOptions.FrameSlides = msoTrue     active.PrintOut() }` |

#### **PrintOptions.HandoutOrder**

返回或设置在打印的讲义上显示幻灯片的页面版式顺序，其中可在一个页面上显示多张幻灯片。可读/写 **PpPrintHandoutOrder** 类型。

**语法**

**express.HandoutOrder**

*express*   一个代表 **PrintOptions** 对象的变量。

**说明**

| PpPrintHandoutOrder 可以是下列 PpPrintHandoutOrder 常量之一。 |
| ------------------------------------------------------------ |
| **ppPrintHandoutHorizontalFirst** 幻灯片水平排序，第一张幻灯片位于左上角，第二张位于其右侧。如果语言设置指定的是从右向左的语言，则第一张幻灯片位于右上角，第二张位于其左侧。 |
| **ppPrintHandoutVerticalFirst** 幻灯片垂直排序，第一张幻灯片位于左上角，第二张位于其下。如果语言设置指定的是从右向左的语言，则第一张幻灯片位于右上角，第二张位于其下。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例设置当前演示文稿的讲义使每个页面包含六张幻灯片，并在讲义上水平排序，然后打印。 function test() {     let active = Application.ActivePresentation     active.PrintOptions.OutputType = ppPrintOutputSixSlideHandouts     active.PrintOptions.HandoutOrder = ppPrintHandoutHorizontalFirst     active.PrintOut() }` |

#### **PrintOptions.NumberOfCopies**

返回或设置要打印的演示文稿份数。默认值为 1。可读/写 **Long** 类型。

**语法**

**express.NumberOfCopies**

*express*   一个代表 **PrintOptions** 对象的变量。

**说明**

指定 **PrintOut**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Presentation/Presentation%20.htm#Presentation.PrintOut)方法的 **Copies** 参数将设置此属性的值。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动演示文稿逐份打印三份。*/ function test() {   let print = Application.ActivePresentation.PrintOptions       print.NumberOfCopies = 3       print.Collate = true       print.Parent.PrintOut() }` |

#### **PrintOptions.OutputType**

返回或设置一个值，该值指示要打印演示文稿的哪些组件（幻灯片、讲义、备注页或大纲）。可读/写 **PpPrintOutputType** 类型。

**语法**

**express.OutputType**

*express*   一个代表 **PrintOptions** 对象的变量。

**说明**

| PpPrintOutputType 可以是下列 PpPrintOutputType 常量之一。 |
| --------------------------------------------------------- |
| **ppPrintOutputBuildSlides**                              |
| **ppPrintOutputFourSlideHandouts**                        |
| **ppPrintOutputNineSlideHandouts**                        |
| **ppPrintOutputNotesPages**                               |
| **ppPrintOutputOneSlideHandouts**                         |
| **ppPrintOutputOutline**                                  |
| **ppPrintOutputSixSlideHandouts**                         |
| **ppPrintOutputSlides**                                   |
| **ppPrintOutputThreeSlideHandouts**                       |
| **ppPrintOutputTwoSlideHandouts**                         |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例打印活动演示文稿的讲义，每页打印六张幻灯片。 function test() {     let active = Application.ActivePresentation     active.PrintOptions.OutputType = ppPrintOutputSixSlideHandouts     active.PrintOut() }` |

#### **PrintOptions.Parent**

返回指定对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **PrintOptions** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。 function test() {     let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes     let addShp = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame     addShp.TextRange.Text = "Test text"     addShp.Parent.Rotation = 45 }` |

#### **PrintOptions.PrintColorType**

返回或设置指定文档的打印方式：黑白方式、纯黑色和纯白色（亦称为高对比度）方式或彩色方式。默认值由打印机设置。可读/写 **PpPrintColorType** 类型。

**语法**

**express.PrintColorType**

*express*   一个代表 **PrintOptions** 对象的变量。

**说明**

| PpPrintColorType 可以是下列 PpPrintColorType 类型常数之一。 |
| ----------------------------------------------------------- |
| **ppPrintBlackAndWhite**                                    |
| **ppPrintColor**                                            |
| **ppPrintPureBlackAndWhite**                                |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `本示例以彩色方式打印活动演示文稿的幻灯片。 function test() {   let appli = Application.ActivePresentation       appli.PrintOptions.PrintColorType = ppPrintColor       appli.PrintOut() }` |

#### **PrintOptions.PrintComments**

设置或返回是否打印批注。可读/写 **MsoTriState** 类型。

**语法**

**express.PrintComments**

*express*   一个代表 **PrintOptions** 对象的变量。

**说明**

| **MsoTriState** 可以是下列 **MsoTriState** 常量之一。 |
| ----------------------------------------------------- |
| **msoCTrue** 不适用于此属性。                         |
| **msoFalse** 默认值。不打印批注。                     |
| **msoTriStateMixed** 不适用于此属性。                 |
| **msoTriStateToggle** 不适用于此属性。                |
| **msoTrue** 打印批注。                                |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例指示 WPP 打印批注。*/ function test(){     Application.ActivePresentation.PrintOptions.PrintComments = msoTrue }` |

#### **PrintOptions.PrintFontsAsGraphics**

决定是否将 TrueType 字体作为图形打印。可读/写 **MsoTriState** 类型。

**语法**

**express.PrintFontsAsGraphics**

*express*   一个代表 **PrintOptions** 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| --------------------------------------------- |
| **msoCTrue**                                  |
| **msoFalse**                                  |
| **msoTriStateMixed**                          |
| **msoTriStateToggle**                         |
| **msoTrue** 将 TrueType 字体作为图形打印。    |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例指定将活动演示文稿中的 TrueType 字体作为图形打印。 Application.ActivePresentation.PrintOptions.PrintFontsAsGraphics = msoTrue` |

#### **PrintOptions.PrintHiddenSlides**

决定是否打印指定演示文稿中的隐藏幻灯片。可读/写 **MsoTriState** 类型。

**语法**

**express.PrintHiddenSlides**

*express*   一个代表 **PrintOptions** 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| --------------------------------------------- |
| **msoCTrue**                                  |
| **msoFalse** 默认值。                         |
| **msoTriStateMixed**                          |
| **msoTriStateToggle**                         |
| **msoTrue** 打印指定演示文稿中的隐藏幻灯片。  |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例打印活动演示文稿的所有幻灯片（不论可见或隐藏）。 function test() {     let active = Application.ActivePresentation     active.PrintOptions.PrintHiddenSlides = msoTrue     active.PrintOut() }` |

#### **PrintOptions.PrintInBackground**

决定是否在后台打印指定的演示文稿。可读/写 **MsoTriState** 类型。

**语法**

**express.PrintInBackground**

*express*   一个代表 **PrintOptions** 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。                |
| ------------------------------------------------------------ |
| **msoCTrue**                                                 |
| **msoFalse**                                                 |
| **msoTriStateMixed**                                         |
| **msoTriStateToggle**                                        |
| **msoTrue** 默认值。在后台打印指定的演示文稿，这意味着可以在打印的同时继续工作。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例在后台打印活动演示文稿。 function test() {     let active = Application.ActivePresentation     active.PrintOptions.PrintInBackground = msoTrue     active.PrintOut() }` |

#### **PrintOptions.Ranges**

返回一个 **PrintRanges** 对象，该对象代表演示文稿中要打印的幻灯片范围。只读。

**语法**

**express.Ranges**

*express*   一个代表 **PrintOptions** 对象的变量。

**说明**

如果不想打印整个演示文稿，必须使用 **Add** 方法为每个要打印的连续的幻灯片组创建一个 **PrintRange** 对象。例如，如果要打印指定演示文稿中的第一张、第三张到第五张以及第八张和第九张幻灯片，则必须创建三个 **PrintRange** 对象：一个代表第一张幻灯片；一个代表第三张到第五张幻灯片；一个代表第八张和第九张幻灯片。有关详细信息，请参阅此属性的示例。

[**RangeType** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/PrintOptions/PrintOptions%20.htm#PrintOptions.RangeType)属性必须设置为 **ppPrintSlideRange** 以应用 **PrintRanges** 集合中的打印范围。

要从 **PrintRanges** 集合中清除现有的所有打印范围，请使用 **ClearAll** 方法。

指定 **PrintOut** 方法的 **To** 和 **From** 参数将设置 **PrintRanges** 对象的内容。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例打印活动演示文稿中的第一张幻灯片、第三张到第五张幻灯片以及第八张和第九张幻灯片。*/ function test() {   let active = Application.ActivePresentation       let print = active.PrintOptions           print.RangeType = ppPrintSlideRange           let ranges = print.Ranges               ranges.Add(1, 1)               ranges.Add(3, 5)               ranges.Add(8, 9)       active.PrintOut() }` |

#### **PrintOptions.RangeType**

返回或设置演示文稿打印范围的类型。可读/写 **PpPrintRangeType** 类型。

**语法**

**express.RangeType**

*express*   一个代表 **PrintOptions** 对象的变量。

**说明**

| PpSlideShowRangeType 可以是下列 PpSlideShowRangeType 常量之一。 |
| ------------------------------------------------------------ |
| **ppShowAll**                                                |
| **ppShowNamedSlideShow**                                     |
| **ppShowSlideRange**                                         |

若要打印在**PrintRanges**集合中定义的幻灯片范围，必须先将 **RangeType** 属性设置为 **ppPrintSlideRange**。将 **RangeType** 设置为除 **ppPrintSlideRange** 之外的值会使在 **PrintRanges** 集合中定义的范围失效。但是这不会以任何方式影响 **PrintRanges** 集合的内容。也就是说，如果定义了一些打印范围，将 **RangeType** 属性设置为除 **ppPrintSlideRange** 之外的值，然后将 **RangeType** 重新设置为 **ppPrintSlideRange**，则之前定义的打印范围将保持不变。

指定 **PrintOut** 方法的 *To* 和 *From* 参数将设置该属性的值。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例打印当前演示文稿的当前幻灯片。*/ function test() {   let active = Application.ActivePresentation       active.PrintOptions.RangeType = ppPrintCurrent       active.PrintOut() }` |

#### **PrintOptions.SlideShowName**

返回或设置自定义幻灯片放映的名称，该幻灯片放映将运行以响应幻灯片放映期间在形状上的鼠标动作（**ActionSetting** 对象）；返回或设置自定义幻灯片放映的名称以便打印（**PrintOptions** 对象）；返回或设置发布为 Web 演示文稿的自定义幻灯片放映的名称（**PublishObject** 对象）。**String** 类型，可读/写。

**语法**

**express.SlideShowName**

*express*   一个代表 **PrintOptions** 对象的变量。

**说明**

必须将 **RangeType** 属性设置为 **ppPrintNamedSlideShow** 才能打印自定义幻灯片放映。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例打印名为“tech talk”的现有自定义幻灯片放映。 function test() {     let print = Application.ActivePresentation.PrintOptions     print.RangeType = Application.Enum.ppPrintNamedSlideShow     print.SlideShowName = "tech talk"     Application.ActivePresentation.PrintOut() }` |

适用环境：web

适用平台：windows/linux