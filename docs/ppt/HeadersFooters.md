**HeadersFooters**



包含指定幻灯片、备注页、讲义或母版上的所有 **HeaderFooter** 对象。

**说明**

每个 **HeaderFooter** 对象代表一个页眉、页脚、日期和时间或幻灯片编号。

**方法**

|                                                              | 名称      | 说明                                                       |
| ------------------------------------------------------------ | --------- | ---------------------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Clear** | 从标尺中清除指定的制表位并将其从 **TabStops** 集合中删除。 |

**属性**

|                                                              | 名称                    | 说明                                                         |
| ------------------------------------------------------------ | ----------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**         | 返回一个 **Application** 对象，该对象表示指定对象的创建者。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DateAndTime**         | 返回一个代表日期和时间项的 **HeaderFooter** 对象，该项将显示在幻灯片的左下角，或者备注页、讲义或大纲的右上角。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayOnTitleSlide** | 决定标题幻灯片上是否显示页脚、日期和时间以及幻灯片编号。可读/写 **MsoTriState** 类型。适用于幻灯片母版。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Footer**              | 返回一个 **HeaderFooter** 对象，该对象代表显示于幻灯片底部或者备注页、讲义或大纲的左下角的页脚。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Header**              | 返回一个 **HeaderFooter** 对象，该对象代表显示于幻灯片顶部或者备注页、讲义或大纲的左下角的页眉。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**              | 返回指定对象的父对象。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SlideNumber**         | 返回一个 **HeaderFooter** 对象，该对象代表位于幻灯片右下角的幻灯片编号，或者位于备注页、打印的讲义或大纲右下角的页码。只读。 |

**成员方法**

#### **HeadersFooters.Clear**

从标尺中清除指定的制表位并将其从 **TabStops** 集合中删除。

**语法**

**express.Clear()**

*express*   一个代表 **HeadersFooters** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例清除当前演示文稿中第一张幻灯片第二个形状中的文本的所有制表位。*/ function test() {     let ts = Application.ActivePresentation.Slides.Item(1).Shapes.Item(2).TextFrame.Ruler.TabStops     for (let i = ts.count; i >= 1; i--) {         ts.Item(i).Clear()     } }` |

**成员属性**

#### **HeadersFooters.Application**

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **HeadersFooters** 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function AddAndSave(pptPre) {     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") }` |

以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes     for (let i = 1; i <= shpOle.Count; i++) {         if (shpOle.Item(i).Type == msoLinkedOLEObject) {             alert(shpOle.Item(i).OLEFormat.Application.Name)         }     } }` |

#### **HeadersFooters.DateAndTime**

返回一个代表日期和时间项的 **HeaderFooter** 对象，该项将显示在幻灯片的左下角，或者备注页、讲义或大纲的右上角。只读。

**语法**

**express.DateAndTime**

*express*   一个代表 **HeadersFooters** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为活动演示文稿的母版幻灯片设置日期和时间格式。该设定将对所有基于该幻灯片母版的幻灯片有效。*/ function test(){ 　　　　let da = Application.ActivePresentation.SlideMaster.HeadersFooters.DateAndTime 　　　　da.Format = ppDateTimeMdyy 　　　　da.UseFormat = true }` |

#### **HeadersFooters.DisplayOnTitleSlide**

决定标题幻灯片上是否显示页脚、日期和时间以及幻灯片编号。可读/写 **MsoTriState** 类型。适用于幻灯片母版。

**语法**

**express.DisplayOnTitleSlide**

*express*   一个代表 **HeadersFooters** 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。                |
| ------------------------------------------------------------ |
| msoCTrue                                                     |
| msoFalse 在除标题幻灯片之外的所有幻灯片上显示页脚、日期和时间以及幻灯片编号。 |
| msoTriStateMixed                                             |
| msoTriStateToggle                                            |
| msoTrue 在标题幻灯片上显示页脚、日期和时间以及幻灯片编号。   |

**示例**

本示例设置标题幻灯片上不显示页脚、日期和时间及幻灯片编号。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActivePresentation.SlideMaster.HeadersFooters.DisplayOnTitleSlide = msoFalse` |

#### **HeadersFooters.Footer**

返回一个 **HeaderFooter** 对象，该对象代表显示于幻灯片底部或者备注页、讲义或大纲的左下角的页脚。只读。

**语法**

**express.Footer**

*express*   一个代表 **HeadersFooters** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为当前演示文稿的幻灯片母版设置脚注的文本，并设定在标题幻灯片上显示脚注、日期和时间及幻灯片编号。*/ function test(){ 　　　　let foot = Application.ActivePresentation.SlideMaster.HeadersFooters 　　　　foot.Footer.Text = "Introduction" 　　　　foot.DisplayOnTitleSlide = true } ` |

#### **HeadersFooters.Header**

返回一个 **HeaderFooter** 对象，该对象代表显示于幻灯片顶部或者备注页、讲义或大纲的左下角的页眉。只读。

**语法**

**express.Header**

*express*   一个代表 **HeadersFooters** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为当前演示文稿的讲义母版设置页眉文本。以大纲或讲义方式打印演示文稿时，该文本将显示于页的左上角。*/ function test(){ 　　　　let myHandHF = Application.ActivePresentation.HandoutMaster.HeadersFooters 　　　　myHandHF.Header.Text = "Third Quarter Report" }` |

#### **HeadersFooters.Parent**

返回指定对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **HeadersFooters** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/ function test() {     let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes     let addShp = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame     addShp.TextRange.Text = "Test text"     addShp.Parent.Rotation = 45 }` |

#### **HeadersFooters.SlideNumber**

返回一个 **HeaderFooter** 对象，该对象代表位于幻灯片右下角的幻灯片编号，或者位于备注页、打印的讲义或大纲右下角的页码。只读。

**语法**

**express.SlideNumber**

*express*   一个代表 **HeadersFooters** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例隐藏当前演示文稿第二张幻灯片的幻灯片编号（如果该编号当前为可见），或显示该幻灯片编号（如果该编号当前为隐藏）。*/ function test() {     let slides = Application.ActivePresentation.Slides.Item(2).HeadersFooters.SlideNumber     if (slides.Visible) {         slides.Visible = false     }     else {         slides.Visible = true     } }` |

适用环境：web

适用平台：windows/linux