**PageSetup**



代表页面设置说明。

**说明**

**PageSetup** 对象包含所有页面设置的属性（左边距、底部边距、纸张大小等）。

使用 **PageSetup** 属性可返回一个 **PageSetup** 对象。下例将打印方向设置为横向，然后打印工作表。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `funtion test(){     let wss = Application.Worksheets.Item("Sheet1")     wss.PageSetup.Orientation = xlLandscape     wss.PrintOut() }` |

**With** 语句使同时设置若干属性变得简单而迅速。下例设置第一张工作表的所有页边距。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){     let psp = Application.Worksheets.Item(1).PageSetup     psp.LeftMargin = Application.InchesToPoints(0.5)     psp.RightMargin = Application.InchesToPoints(0.75)     psp.TopMargin = Application.InchesToPoints(1.5)     psp.BottomMargin = Application.InchesToPoints(1)     psp.HeaderMargin = Application.InchesToPoints(0.5)     psp.FooterMargin = Application.InchesToPoints(0.5) }` |

**方法**

|                                                              | 名称             | 说明                                           |
| ------------------------------------------------------------ | ---------------- | ---------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PrintQuality** | 返回或设置打印质量。**Variant** 类型，可读写。 |

**属性**

|                                                              | 名称                               | 说明                                                         |
| ------------------------------------------------------------ | ---------------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AlignMarginsHeaderFooter**       | 如果 ET 以页面设置选项中设置的边距对齐页眉和页脚，则返回 **True**。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**                    | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BlackAndWhite**                  | 如果指定文档中的元素以黑白方式打印，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BottomMargin**                   | 以磅为单位返回或设置底端边距的大小。**Double** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CenterFooter**                   | 居中对齐 **PageSetup** 对象中的页脚信息。可读/写 **String** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CenterFooterPicture**            | 返回一个 **Graphic** 对象，该对象表示页脚中间部分的图片。用于设置图片的属性。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CenterHeader**                   | 居中对齐 **PageSetup** 对象中的页眉信息。可读/写 **String** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CenterHeaderPicture**            | 返回一个 **Graphic** 对象，该对象表示页眉中间部分的图片。用于设置图片的属性。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CenterHorizontally**             | 如果在页面的水平居中位置打印指定工作表，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CenterVertically**               | 如果在页面的垂直居中位置打印指定工作表，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                        | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DifferentFirstPageHeaderFooter** | 如果在第一页使用不同的页眉或页脚，则为 **True**。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Draft**                          | 如果打印工作表时不打印其中的图形，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EvenPage**                       | 返回或设置工作簿或节的偶数页上的文本对齐方式。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FirstPage**                      | 返回或设置工作簿或节的第一页上的文本对齐方式。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FirstPageNumber**                | 返回或设置打印指定工作表时第一页的页号。如果设为 **xlAutomatic**，则 ET 采用第一页的页号。默认值为 **xlAutomatic**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FitToPagesTall**                 | 返回或设置打印工作表时，对工作表进行缩放使用的页高。仅应用于工作表。**Variant** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FitToPagesWide**                 | 返回或设置打印工作表时，对工作表进行缩放使用的页宽。仅应用于工作表。**Variant** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FooterMargin**                   | 以磅为单位返回或设置页脚到页面底端的距离。**Double** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HeaderMargin**                   | 以磅为单位返回或设置页面顶端到页眉的距离。**Double** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LeftFooter**                     | 返回或设置工作簿或节的左页脚上的文本对齐方式。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LeftFooterPicture**              | 返回一个 **Graphic** 对象，该对象表示页脚左边的图片。用于设置图片的属性。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LeftHeader**                     | 返回或设置工作簿或节的左页眉上的文本对齐方式。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LeftHeaderPicture**              | 返回一个 **Graphic** 对象，该对象表示页眉左边的图片。用于设置图片的属性。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LeftMargin**                     | 以磅为单位返回或设置左边距的大小。**Double** 类型，可读写。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OddAndEvenPagesHeaderFooter**    | 如果指定的 **PageSetup** 对象的奇数页和偶数页具有不同的页眉和页脚，则为 **True**。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Order**                          | 返回或设置一个 **XlOrder** 值，它代表 ET 打印一张大工作表时所使用的页编号的次序。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Orientation**                    | 返回或设置一个 **XlPageOrientation** 值，它代表纵向或横向打印模式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Pages**                          | 返回或设置 **Pages** 集合中的页数。                          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PaperSize**                      | 返回或设置纸张的大小。XlPaperSize 类型，可读写。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                         | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PrintArea**                      | 以字符串返回或设置要打印的区域，该字符串使用宏语言的 A1 样式的引用。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PrintComments**                  | 返回或设置批注随工作表打印的方式。**XlPrintLocation** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PrintErrors**                    | 设置或返回一个 **XlPrintErrors** 常量，该常量指定显示的打印错误类型。该功能允许用户在打印工作表时取消错误显示。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PrintGridlines**                 | 如果在页面上打印单元格网格线，则该值为 **True**。仅应用于工作表。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PrintHeadings**                  | 如果打印本页时同时打印行标题和列标题，则该值为 **True**。仅应用于工作表。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PrintNotes**                     | 如果打印工作表时单元格批注作为尾注一起打印，则该值为 **True**。仅应用于工作表。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PrintTitleColumns**              | 返回或设置包含在每一页的左边重复出现的单元格的列，用宏语言 A1-样式中的字符串表示。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PrintTitleRows**                 | 返回或设置那些包含在每一页顶部重复出现的单元格的行，用宏语言字符串以 A1 样式表示法表示。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RightFooter**                    | 返回或设置页面右边缘与页脚右边界之间的距离（以磅为单位）。可读/写 **String** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RightFooterPicture**             | 返回一个 **Graphic** 对象，该对象代表页脚右边的图片，用于设置图片的属性。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RightHeader**                    | 返回或设置页眉的右边部分内容。可读/写 **String** 类型。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RightHeaderPicture**             | 指定右页眉中应显示的图形图像。只读。                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RightMargin**                    | 以磅为单位返回或设置右边距的大小。**Double** 类型，可读写。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ScaleWithDocHeaderFooter**       | 返回或设置页眉和页脚是否在文档大小更改时随文档缩放。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TopMargin**                      | 以磅为单位返回或设置上边距的大小。**Double** 类型，可读写。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Zoom**                           | 返回或设置一个 **Variant** 值，它代表一个数值在 10% 到 400% 之间的百分比，该百分比为 ET 打印工作表时的缩放比例。 |

**成员方法**

#### **PageSetup.PrintQuality**

返回或设置打印质量。**Variant** 类型，可读写。

**语法**

**express.PrintQuality(Index)**

*express*   一个代表 **PageSetup** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Index*  | 可选          | **Variant**  | 水平方向打印质量 (1) 或垂直方向打印质量 (2)。有些打印机可能不支持垂直方向打印质量设置。如果不指定该参数，则 PrintQuality 属性将返回（或可被设置为）一个二元数组，该数组包含水平方向和垂直方向的打印质量。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   /*用非正方形像素设置打印机上的打印质量。数组指定水平方向和垂直方向的打印质量。本示例是否出现错误取决于使用的打印机驱动程序。*/   Application.Worksheets.Item("Sheet1").PageSetup.PrintQuality = [240, 140]      /*显示水平方向打印质量的当前设置。*/   alert("Horizontal Print Quality is " + Application.Worksheets.Item("Sheet1").PageSetup.PrintQuality(1)) }` |

**成员属性**

#### **PageSetup.AlignMarginsHeaderFooter**

如果 ET 以页面设置选项中设置的边距对齐页眉和页脚，则返回 **True**。可读/写 **Boolean** 类型。

**语法**

**express.AlignMarginsHeaderFooter**

*express*   一个代表 **PageSetup** 对象的变量。

#### **PageSetup.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **PageSetup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   /*本示例显示一条有关创建 myObject 的应用程序的消息。*/   let myObject = Application.ActiveWorkbook   if (myObject.Application.Value == "ET"){       alert("This is an ET Application object.")   }   else{       alert("This is not an ET Application object.")   } }` |

#### **PageSetup.BlackAndWhite**

如果指定文档中的元素以黑白方式打印，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.BlackAndWhite**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

该属性仅适用于工作表页面。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置工作表 Sheet1 以黑白方式打印。*/ Application.Worksheets.Item("Sheet1").PageSetup.BlackAndWhite = true` |

#### **PageSetup.BottomMargin**

以磅为单位返回或设置底端边距的大小。**Double** 类型，可读写。

**语法**

**express.BottomMargin**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

边距的设置和返回均以磅为单位。可使用 **InchesToPoints** 方法进行英寸到磅值的转换，也可使用 **CentimetersToPoints** 方法进行厘米到磅值的转换。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){     /*将工作表 Sheet1 的底端边距设为 0.5 英寸（36 磅）。*/   Application.Worksheets.Item("Sheet1").PageSetup.BottomMargin = Application.InchesToPoints(0.5)   Application.Worksheets.Item("Sheet1").PageSetup.BottomMargin = 36      /*显示 Sheet1 底端边距的当前设定值。*/   let marginInches = Application.Worksheets.Item("Sheet1").PageSetup.BottomMargin /    Application.InchesToPoints(1)   alert("The current bottom margin is " + marginInches + " inches") }` |

#### **PageSetup.CenterFooter**

居中对齐 **PageSetup** 对象中的页脚信息。可读/写 **String** 类型。

**语法**

**express.CenterFooter**

*express*   一个代表 **PageSetup** 对象的变量。

#### **PageSetup.CenterFooterPicture**

返回一个 **Graphic** 对象，该对象表示页脚中间部分的图片。用于设置图片的属性。

**语法**

**express.CenterFooterPicture**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

**CenterFooterPicture** 属性为只读，但并非它的所有属性都为只读。

要求“&G”为 **CenterFooter** 属性字符串的组成部分，以便在页脚的中间显示图像。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| ` function InsertPicture() {   /*本示例将 C:\ 驱动器中名为“Sample.jpg”的图片添加到页脚的中间部分(“Sample.jpg”文件位于 C:\ 驱动器上。)*/   let picture = Application.ActiveSheet.PageSetup.CentertFooterPicture   picture.FileName = "C:\\Sample.jpg"   picture.Height = 275.25   picture.Width = 463.5   picture.Brightness = 0.36   picture.ColorType = msoPictureGrayscale   picture.Contrast = 0.39   picture.CropBottom = -14.4   picture.CropLeft = -28.8   picture.CropRight = -14.4   picture.CropTop = 21.6   									   // Enable the image to show up in the center footer.   ActiveSheet.PageSetup.CenterFooter = "&G"     }` |

#### **PageSetup.CenterHeader**

居中对齐 **PageSetup** 对象中的页眉信息。可读/写 **String** 类型。

**语法**

**express.CenterHeader**

*express*   一个代表 **PageSetup** 对象的变量。

#### **PageSetup.CenterHeaderPicture**

返回一个 **Graphic** 对象，该对象表示页眉中间部分的图片。用于设置图片的属性。

**语法**

**express.CenterHeaderPicture**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

**CenterHeaderPicture** 属性为只读，但并非它的所有属性都为只读。

要求“&G”为 **CenterHeader** 属性字符串的组成部分，以便在中间的页眉中显示图像。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| ` function InsertPicture() {   /*本示例将 C:\ 驱动器中名为“Sample.jpg”的图片添加到页眉的中间部分(“Sample.jpg”文件位于 C:\ 驱动器上。)*/   let picture = Application.ActiveSheet.PageSetup.CentertHeaderPicture   picture.FileName = "C:\\Sample.jpg"   picture.Height = 275.25   picture.Width = 463.5   picture.Brightness = 0.36   picture.ColorType = msoPictureGrayscale   picture.Contrast = 0.39   picture.CropBottom = -14.4   picture.CropLeft = -28.8   picture.CropRight = -14.4   picture.CropTop = 21.6   										   // Enable the image to show up in the center footer.   Application.ActiveSheet.PageSetup.CenterFooter = "&G"    										 }` |

#### **PageSetup.CenterHorizontally**

如果在页面的水平居中位置打印指定工作表，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.CenterHorizontally**

*express*   一个代表 **PageSetup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将工作表 Sheet1 设置成水平居中打印。*/ Application.Worksheets.Item("Sheet1").PageSetup.CenterHorizontally = true` |

#### **PageSetup.CenterVertically**

如果在页面的垂直居中位置打印指定工作表，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.CenterVertically**

*express*   一个代表 **PageSetup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 Sheet1 设为垂直居中打印。*/ Application.Worksheets.Item("Sheet1").PageSetup.CenterVertically = true` |

#### **PageSetup.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **PageSetup.DifferentFirstPageHeaderFooter**

如果在第一页使用不同的页眉或页脚，则为 **True**。可读/写 **Boolean** 类型。

**语法**

**express.DifferentFirstPageHeaderFooter**

*express*   一个代表 **PageSetup** 对象的变量。

#### **PageSetup.Draft**

如果打印工作表时不打印其中的图形，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.Draft**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

将该属性设置为 **True** 可加快打印速度（但是不打印其中的图形）。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例关闭对 Sheet1 中图形的打印。*/ Application.Worksheets.Item("Sheet1").PageSetup.Draft = true` |

#### **PageSetup.EvenPage**

返回或设置工作簿或节的偶数页上的文本对齐方式。

**语法**

**express.EvenPage**

*express*   一个代表 **PageSetup** 对象的变量。

#### **PageSetup.FirstPage**

返回或设置工作簿或节的第一页上的文本对齐方式。

**语法**

**express.FirstPage**

*express*   一个代表 **PageSetup** 对象的变量。

#### **PageSetup.FirstPageNumber**

返回或设置打印指定工作表时第一页的页号。如果设为 **xlAutomatic**，则 ET 采用第一页的页号。默认值为 **xlAutomatic**。**Long** 类型，可读写。

**语法**

**express.FirstPageNumber**

*express*   一个代表 **PageSetup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 Sheet1 打印时的第一页的页号设置为 100。*/ Application.Worksheets.Item("Sheet1").PageSetup.FirstPageNumber = 100` |

#### **PageSetup.FitToPagesTall**

返回或设置打印工作表时，对工作表进行缩放使用的页高。仅应用于工作表。**Variant** 类型，可读写。

**语法**

**express.FitToPagesTall**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

如果该属性值为 **False**，ET 将根据 **FitToPagesWide** 属性缩放工作表。

如果 **Zoom** 属性值为 **True**，将忽略 **FitToPagesTall** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){     /*本示例设置 ET 准确按照一页的宽度和高度打印 Sheet1。*/   let pagesetup = Application.Worksheets.Item("Sheet1").PageSetup   pagesetup.Zoom = false   pagesetup.FitToPagesTall = 1   pagesetup.FitToPagesWide = 1 }` |

#### **PageSetup.FitToPagesWide**

返回或设置打印工作表时，对工作表进行缩放使用的页宽。仅应用于工作表。**Variant** 类型，可读写。

**语法**

**express.FitToPagesWide**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

如果该属性值为 **False**，ET 将根据 **FitToPagesTall** 属性缩放工作表。

如果 **Zoom** 属性值为 **True**，则忽略 **FitToPagesWide** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   /*本示例设置 ET 准确按照一页的宽度和高度打印 Sheet1。*/   let pagesetup = Application.Worksheets.Item("Sheet1").PageSetup   pagesetup.Zoom = false   pagesetup.FitToPagesTall = 1   pagesetup.FitToPagesWide = 1 }` |

#### **PageSetup.FooterMargin**

以磅为单位返回或设置页脚到页面底端的距离。**Double** 类型，可读写。

**语法**

**express.FooterMargin**

*express*   一个代表 **PageSetup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 Sheet1 的页脚边矩设置为 0.5 英寸。*/ Application.Worksheets.Item("Sheet1").PageSetup.FooterMargin = Application.InchesToPoints(0.5)` |

#### **PageSetup.HeaderMargin**

以磅为单位返回或设置页面顶端到页眉的距离。**Double** 类型，可读写。

**语法**

**express.HeaderMargin**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

边距的设置和返回均以磅为单位。可使用 **InchesToPoints** 方法进行英寸到磅值的转换，也可使用 **CentimetersToPoints** 方法进行厘米到磅值的转换。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 Sheet1 的页眉边距设置为 0.5 英寸。*/ Application.Worksheets.Item("Sheet1").PageSetup.HeaderMargin = Application.InchesToPoints(0.5)` |

#### **PageSetup.LeftFooter**

返回或设置工作簿或节的左页脚上的文本对齐方式。

**语法**

**express.LeftFooter**

*express*   一个代表 **PageSetup** 对象的变量。

#### **PageSetup.LeftFooterPicture**

返回一个 **Graphic** 对象，该对象表示页脚左边的图片。用于设置图片的属性。

**语法**

**express.LeftFooterPicture**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

**LeftFooterPicture** 属性为只读，但并非其所有属性都为只读。

| ![img](https://qn.cache.wpscdn.cn/gif/close.gif)注释         |
| ------------------------------------------------------------ |
| 要求 "&G" 为 **LeftFooter** 属性字符串的组成部分，以便在左边的页脚中显示图像。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function InsertPicture() {   /*本示例将 C:\ 驱动器中名为“Sample.jpg”的图片添加到页脚的左边(假定“Sample.jpg”文件位于 C:\ 驱动器上。)*/   let picture = Application.ActiveSheet.PageSetup.LeftFooterPicture   picture.FileName = "C:\\Sample.jpg"   picture.Height = 275.25   picture.Width = 463.5   picture.Brightness = 0.36   picture.ColorType = msoPictureGrayscale   picture.Contrast = 0.39   picture.CropBottom = -14.4   picture.CropLeft = -28.8   picture.CropRight = -14.4   picture.CropTop = 21.6   										   // Enable the image to show up in the center footer.   Application.ActiveSheet.PageSetup.CenterFooter = "&G"    										 }` |

#### **PageSetup.LeftHeader**

返回或设置工作簿或节的左页眉上的文本对齐方式。

**语法**

**express.LeftHeader**

*express*   一个代表 **PageSetup** 对象的变量。

#### **PageSetup.LeftHeaderPicture**

返回一个 **Graphic** 对象，该对象表示页眉左边的图片。用于设置图片的属性。

**语法**

**express.LeftHeaderPicture**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

**LeftHeaderPicture** 属性为只读，但并非其所有属性都为只读。

| ![img](https://qn.cache.wpscdn.cn/gif/close.gif)注释         |
| ------------------------------------------------------------ |
| 要求 "&G" 为 **LeftHeader** 属性字符串的组成部分，以便在左边的页眉中显示图像。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {   /*本示例将 C: 驱动器中名为“Sample.jpg”的图片添加到页眉的左边。(假定“Sample.jpg”文件位于 C: 驱动器上。)*/   let picture = Application.ActiveSheet.PageSetup.LeftHeaderPicture   picture.FileName = "C:\\Sample.jpg"   picture.Height = 275.25   picture.Width = 463.5   picture.Brightness = 0.36   picture.ColorType = msoPictureGrayscale   picture.Contrast = 0.39   picture.CropBottom = -14.4   picture.CropLeft = -28.8   picture.CropRight = -14.4   picture.CropTop = 21.6   										   // Enable the image to show up in the center footer.   Application.ActiveSheet.PageSetup.CenterFooter = "&G"    										 }` |

#### **PageSetup.LeftMargin**

以磅为单位返回或设置左边距的大小。**Double** 类型，可读写。

**语法**

**express.LeftMargin**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

边距的设置和返回均以磅为单位。可使用 **InchesToPoints** 方法进行英寸到磅值的转换，也可使用 **CentimetersToPoints** 方法进行厘米到磅值的转换。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){    /*将 Sheet1 的左边距设置为 1.5 英寸。*/   Application.Worksheets.Item("Sheet1").PageSetup.LeftMargin = Application.InchesToPoints(1.5)      /*将 Sheet1 的左页边距设置为 2 厘米。*/   Application.Worksheets.Item("Sheet1").PageSetup.LeftMargin = Application.CentimetersToPoints(2)      /*显示 Sheet1 的左边距的当前设定值。*/   let marginInches = Application.Worksheets.Item("Sheet1").PageSetup.LeftMargin /   Application.InchesToPoints(1)   alert("The current left margin is " + marginInches + " inches") }` |

#### **PageSetup.OddAndEvenPagesHeaderFooter**

如果指定的 **PageSetup** 对象的奇数页和偶数页具有不同的页眉和页脚，则为 **True**。可读/写 **Boolean** 类型。

**语法**

**express.OddAndEvenPagesHeaderFooter**

*express*   一个代表 **PageSetup** 对象的变量。

#### **PageSetup.Order**

返回或设置一个 **XlOrder** 值，它代表 ET 打印一张大工作表时所使用的页编号的次序。

**语法**

**express.Order**

*express*   一个代表 **PageSetup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置打印 Sheet1 时将其分为多页打印。按照从左到右，从上到下的顺序编号和打印。*/ Application.Worksheets.Item("Sheet1").PageSetup.Order = xlOverThenDown` |

#### **PageSetup.Orientation**

返回或设置一个 **XlPageOrientation** 值，它代表纵向或横向打印模式。

**语法**

**express.Orientation**

*express*   一个代表 **PageSetup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 Sheet1 设置为横向打印。*/ Application.Worksheets.Item("Sheet1").PageSetup.Orientation = xlLandscape` |

#### **PageSetup.Pages**

返回或设置 **Pages** 集合中的页数。

**语法**

**express.Pages**

*express*   一个代表 **PageSetup** 对象的变量。

#### **PageSetup.PaperSize**

返回或设置纸张的大小。XlPaperSize 类型，可读写。

**语法**

**express.PaperSize**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

| **XlPaperSize** 可为以下 **XlPaperSize** 常量之一。          |
| ------------------------------------------------------------ |
| **xlPaper11x17**：11 in. x 17 in.                            |
| **xlPaperA4**：A4 (210 mm x 297 mm)                          |
| **xlPaperA5**：A5 (148 mm x 210 mm)                          |
| **xlPaperB5**：A5 (148 mm x 210 mm)                          |
| **xlPaperDsheet**：D 型纸                                    |
| **xlPaperEnvelope11**：信封 #11 (4-1/2 in. x 10-3/8 in.)     |
| **xlPaperEnvelope14**：信封 #14 (5 in. x 11-1/2 in.)         |
| **xlPaperEnvelopeB4**：信封 B4 (250 mm x 353 mm)             |
| **xlPaperEnvelopeB6**：信封 (176 mm x 125 mm)                |
| **xlPaperEnvelopeC4**：信封 C4 (229 mm x 324 mm)             |
| **xlPaperEnvelopeC6**：信封 C6 (114 mm x 162 mm)             |
| **xlPaperEnvelopeDL**：信封 DL (110 mm x 220 mm)             |
| **xlPaperEnvelopeMonarch**：君主式信封 (3-7/8 in. x 7-1/2 in.) |
| **xlPaperEsheet**：E 型纸                                    |
| **xlPaperFanfoldLegalGerman**：德国正规复写簿 (8-1/2 in. x 13 in.) |
| **xlPaperFanfoldUS**：美国标准复写簿 (14-7/8 in. x 11 in.)   |
| **xlPaperLedger**：帐单 (17 in. x 11 in.)                    |
| **xlPaperLetter**：信函 (8-1/2 in. x 11 in.)                 |
| **xlPaperNote**：便笺 (8-1/2 in. x 11 in.)                   |
| **xlPaperStatement**：报告单 (5-1/2 in. x 8-1/2 in.)         |
| **xlPaperUser**：用户自定义                                  |
| **xlPaper10x14**：10 in. x 14 in.                            |
| **xlPaperA3**：A3 (297 mm x 420 mm)                          |
| **xlPaperA4Small**：A4（小）(210 mm x 297 mm)                |
| **xlPaperB4**：B4 (250 mm x 354 mm)                          |
| **xlPaperCsheet**：C 型纸                                    |
| **xlPaperEnvelope10**：信封 #10 (4-1/8 in. x 9-1/2 in.)      |
| **xlPaperEnvelope12**：信封 #12 (4-1/2 in. x 11 in.)         |
| **xlPaperEnvelope9**：信封 #9 (3-7/8 in. x 8-7/8 in.)        |
| **xlPaperEnvelopeB5**：信封 B5 (176 mm x 250 mm)             |
| **xlPaperEnvelopeC3**：信封 C3 (324 mm x 458 mm)             |
| **xlPaperEnvelopeC5**：信封 C5 (162 mm x 229 mm)             |
| **xlPaperEnvelopeC65**：信封 C65 (114 mm x 229 mm)           |
| **xlPaperEnvelopeItaly**：信封 (110 mm x 230 mm)             |
| **xlPaperEnvelopePersonal**：信封 (3-5/8 in. x 6-1/2 in.)    |
| **xlPaperExecutive**：行政公文纸 (7-1/2 in. x 10-1/2 in.)    |
| **xlPaperFanfoldStdGerman**：德国正规复写簿 (8-1/2 in. x 13 in.) |
| **xlPaperFolio**：对开 (8-1/2 in. x 13 in.)                  |
| **xlPaperLegal**：正式信封 (8-1/2 in. x 14 in.)              |
| **xlPaperLetterSmall**：简式信纸 (8-1/2 in. x 11 in.)        |
| **xlPaperQuarto**：四开 (215 mm x 275 mm)                    |
| **xlPaperTabloid**：文摘 (11 in. x 17 in.)                   |

| ![img](https://qn.cache.wpscdn.cn/gif/close.gif)注释 |
| ---------------------------------------------------- |
| 上面这些纸张大小可能不为所有打印机支持。             |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 Sheet1 的纸张大小设置为正式信封。*/ Application.Worksheets.Item("Sheet1").PageSetup.PaperSize = xlPaperLegal` |

#### **PageSetup.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **PageSetup** 对象的变量。

#### **PageSetup.PrintArea**

以字符串返回或设置要打印的区域，该字符串使用宏语言的 A1 样式的引用。**String** 类型，可读写。

**语法**

**express.PrintArea**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

将该属性设置为 **False** 或空字符串 ("")，可打印整个工作表。

该属性仅适用于工作表页面。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   /*将打印区域设置为 Sheet1 上的单元格 A1:C5。*/   Application.Worksheets.Item("Sheet1").PageSetup.PrintArea = "$A$1:$C$5"      /*将 sheet1 的打印区域设置为当前区。注意使用 Address 属性返回 A1-样式的地址。*/   Application.Worksheets.Item("Sheet1").Activate()   Application.ActiveSheet.PageSetup.PrintArea = ActiveCell.CurrentRegion.Address }` |

#### **PageSetup.PrintComments**

返回或设置批注随工作表打印的方式。**XlPrintLocation** 类型，可读写。

**语法**

**express.PrintComments**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

| **XlPrintLocation** 可为以下 **XlPrintLocation** 常量之一。 |
| ----------------------------------------------------------- |
| **xlPrintInPlace**                                          |
| **xlPrintNoComments**                                       |
| **xlPrintSheetEnd**                                         |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*打印第一张工作表时，本示例使批注以尾注的方式打印。*/ Application.Worksheets.Item(1).PageSetup.PrintComments = xlPrintSheetEnd` |

#### **PageSetup.PrintErrors**

设置或返回一个 **XlPrintErrors** 常量，该常量指定显示的打印错误类型。该功能允许用户在打印工作表时取消错误显示。可读写。

**语法**

**express.PrintErrors**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

| **XlPrintErrors** 可为以下 **XlPrintErrors** 常量之一。 |
| ------------------------------------------------------- |
| **xlPrintErrorsBlank**                                  |
| **xlPrintErrorsDash**                                   |
| **xlPrintErrorsDisplayed**                              |
| **xlPrintErrorsNA**                                     |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function UsePrintErrors() {     /*在本示例中，ET 在活动工作表中使用返回错误的公式。设置 PrintErrors 属性以显示短划线。打印预览窗口显示打印错误的短划线。(假定已安装了打印驱动器)*/     let wksOne = Application.ActiveSheet 									     // Create a formula that returns an error value.     Range("A1").Value = 1     Range("A2").Value = 0     Range("A3").Formula = "=A1/A2" 									     // Change print errors to display dashes.     wksOne.PageSetup.PrintErrors = xlPrintErrorsDash 									     // Use the Print Preview window to see the dashes used for print errors.     ActiveWindow.SelectedSheets.PrintPreview() }` |

#### **PageSetup.PrintGridlines**

如果在页面上打印单元格网格线，则该值为 **True**。仅应用于工作表。**Boolean** 类型，可读写。

**语法**

**express.PrintGridlines**

*express*   一个代表 **PageSetup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在打印 Sheet1 时同时打印单元格网格线。*/ Application.Worksheets.Item("Sheet1").PageSetup.PrintGridlines = true` |

#### **PageSetup.PrintHeadings**

如果打印本页时同时打印行标题和列标题，则该值为 **True**。仅应用于工作表。**Boolean** 类型，可读写。

**语法**

**express.PrintHeadings**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

**DisplayHeadings** 属性控制标题在屏幕上的显示。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例关闭打印 Sheet1 中的标题。*/ Application.Worksheets.Item("Sheet1").PageSetup.PrintHeadings = false` |

#### **PageSetup.PrintNotes**

如果打印工作表时单元格批注作为尾注一起打印，则该值为 **True**。仅应用于工作表。**Boolean** 类型，可读写。

**语法**

**express.PrintNotes**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

使用 **PrintComments** 属性可以文本框或尾注形式打印批注。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例关闭打印批注。*/ Application.Worksheets.Item("Sheet1").PageSetup.PrintNotes = false` |

#### **PageSetup.PrintTitleColumns**

返回或设置包含在每一页的左边重复出现的单元格的列，用宏语言 A1-样式中的字符串表示。**String** 类型，可读写。

**语法**

**express.PrintTitleColumns**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

如果仅指定列的一部分，ET 将自动把该区域扩展为整个列。

将该属性设置为 **False** 或空字符串 ("")，将会关闭标题列。

该属性仅适用于工作表页面。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   /*本示例将第三行定义为标题行，将第一列到第三列定义为标题列。*/   Application.Worksheets.Item("Sheet1").Activate()   Application.ActiveSheet.PageSetup.PrintTitleRows = ActiveSheet.Rows.Item(3).Address   Application.ActiveSheet.PageSetup.PrintTitleColumns = ActiveSheet.Columns.Item("A:C").Address }` |

#### **PageSetup.PrintTitleRows**

返回或设置那些包含在每一页顶部重复出现的单元格的行，用宏语言字符串以 A1 样式表示法表示。**String** 类型，可读写。

**语法**

**express.PrintTitleRows**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

如果仅指定行的一部分，ET 将把该区域扩展为整个行。

将该属性设置为 **False** 或空字符串 ("")，将会关闭标题行。

该属性仅适用于工作表页面。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){     /*本示例将第三行定义为标题行，将第一列到第三列定义为标题列。*/   Application.Worksheets.Item("Sheet1").Activate()   Application.ActiveSheet.PageSetup.PrintTitleRows = ActiveSheet.Rows.Item(3).Address   Application.ActiveSheet.PageSetup.PrintTitleColumns = ActiveSheet.Columns.Item("A:C").Address }` |

#### **PageSetup.RightFooter**

返回或设置页面右边缘与页脚右边界之间的距离（以磅为单位）。可读/写 **String** 类型。

**语法**

**express.RightFooter**

*express*   一个代表 **PageSetup** 对象的变量。

#### **PageSetup.RightFooterPicture**

返回一个 **Graphic** 对象，该对象代表页脚右边的图片，用于设置图片的属性。

**语法**

**express.RightFooterPicture**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

**RightFooterPicture** 属性为只读，但并非其所有属性都为只读。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function InsertPicture() {     /*本示例将 C: 驱动器中名为“Sample.jpg”的图片添加到页脚的右边。（假定“Sample.jpg”文件位于 C: 驱动器上。）*/     et picture = Application.ActiveSheet.PageSetup.RightFooterPicture     picture.FileName = "C:\\Sample.jpg"     picture.Height = 275.25     picture.Width = 463.5     picture.Brightness = 0.36     picture.ColorType = msoPictureGrayscale     picture.Contrast = 0.39     picture.CropBottom = -14.4     picture.CropLeft = -28.8     picture.CropRight = -14.4     picture.CropTop = 21.6 										     // Enable the image to show up in the center footer.     Application.ActiveSheet.PageSetup.CenterFooter = "&G" 										 }` |

#### **PageSetup.RightHeader**

返回或设置页眉的右边部分内容。可读/写 **String** 类型。

**语法**

**express.RightHeader**

*express*   一个代表 **PageSetup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在每页的右上角打印文件名。*/ Application.Worksheets.Item("Sheet1").PageSetup.RightHeader = "&F"` |

#### **PageSetup.RightHeaderPicture**

指定右页眉中应显示的图形图像。只读。

**语法**

**express.RightHeaderPicture**

*express*   一个代表 **PageSetup** 对象的变量。

#### **PageSetup.RightMargin**

以磅为单位返回或设置右边距的大小。**Double** 类型，可读写。

**语法**

**express.RightMargin**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

边距的设置和返回均以磅为单位。可使用 **InchesToPoints** 方法进行英寸到磅值的转换，也可使用 **CentimetersToPoints** 方法进行厘米到磅值的转换。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   /*将 Sheet1 的右边距设置为 1.5 英寸。*/   Application.Worksheets.Item("Sheet1").PageSetup.RightMargin = Application.InchesToPoints(1.5)      /*将 Sheet1 的右边距设置为 2 厘米。*/   Application.Worksheets.Item("Sheet1").PageSetup.RightMargin = Application.CentimetersToPoints(2)      /*显示 Sheet1 中右边距的当前设置。*/   let marginInches = Application.Worksheets.Item("Sheet1").PageSetup.RightMargin /   Application.InchesToPoints(1)   alert("The current right margin is " + marginInches + " inches") }` |

#### **PageSetup.ScaleWithDocHeaderFooter**

返回或设置页眉和页脚是否在文档大小更改时随文档缩放。可读/写 **Boolean** 类型。

**语法**

**express.ScaleWithDocHeaderFooter**

*express*   一个代表 **PageSetup** 对象的变量。

#### **PageSetup.TopMargin**

以磅为单位返回或设置上边距的大小。**Double** 类型，可读写。

**语法**

**express.TopMargin**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

边距的设置和返回均以磅为单位。可使用 **InchesToPoints** 方法进行英寸到磅值的转换，也可使用 **CentimetersToPoints** 方法进行厘米到磅值的转换。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   /*将 Sheet1 的上边距设置为 0.5 英寸（36 磅）*/   Application.Worksheets.Item("Sheet1").PageSetup.TopMargin = Application.InchesToPoints(0.5)   Application.Worksheets.Item("Sheet1").PageSetup.TopMargin = 36      /*本示例显示上边距的当前设置值。*/   let marginInches = Application.ActiveSheet.PageSetup.TopMargin /   Application.InchesToPoints(1)   alert("The current top margin is " + marginInches + " inches") }` |

#### **PageSetup.Zoom**

返回或设置一个 **Variant** 值，它代表一个数值在 10% 到 400% 之间的百分比，该百分比为 ET 打印工作表时的缩放比例。

**语法**

**express.Zoom**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

此属性仅适用于工作表。

如果此属性设置为 **False**，则 **FitToPagesWide** 和 **FitToPagesTall** 属性控制工作表缩放的方式。

所有缩放均保持原文档的长宽比例。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例设置打印 Sheet1 时的缩放比例为 150%。*/ Application.Worksheets.Item("Sheet1").PageSetup.Zoom = 150` |

适用环境：web

适用平台：windows/linux