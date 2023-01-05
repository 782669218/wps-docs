**TextRange**



包含附加到形状的文本，以及用于操作文本的属性和方法。

**说明**

以下示例说明如何执行下列操作：

- 返回任意指定形状中的文本范围。
- 返回选定范围中的文本范围。
- 返回文本范围中的特定字符、单词、行、句子或段落。
- 查找和替换文本范围内的文本。
- 向文本范围中插入文本、日期和时间或幻灯片编号。
- 将光标定位到文本范围内所需的任意位置。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*使用 TextFrame 对象的 TextRange 属性返回任意指定形状的 TextRange 对象。使用 Text 属性返回 TextRange 对象中的文本字符串。以下示例向 myDocument 中添加一个矩形并设置其包含的文本。*/ function test() {     let myDocument = Application.ActivePresentation.Slides.Item(1)     myDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame.TextRange.Text = "Here is some test text" }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*因为 Text 属性是 TextRange 对象的默认属性，所以以下两个语句是等效的。*/ Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.Text = "Here is some test text" Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange = "Here is some test text"` |

 

使用 

HasTextFrame

 属性判断形状是否含有文本框，然后使用 

HasText

 属性判断该文本框是否包含文本。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*使用 Selection 对象的 TextRange 属性返回当前选定的文字。以下示例将选定内容复制到剪贴板。*/ Application.ActiveWindow.Selection.TextRange.Copy()` |

 

使用下列方法之一可返回 

TextRange

 对象中的部分文本： 

[Characters](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/TextRange/TextRange%20.htm#TextRange.Characters)

、 

Lines

、 

[Paragraphs](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/TextRange/TextRange%20.htm#TextRange.Paragraphs)

、 

Runs

、 

Sentences

 或 

Words

。

使用 **Find** 和 **Replace** 方法可查找和替换文本范围内的文本。

使用下列方法之一可向 **TextRange** 对象中插入字符：**InsertAfter**、**InsertBefore**、**InsertDateTime**、[**InsertSlideNumber** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/TextRange/TextRange%20.htm#TextRange.InsertSlideNumber)或 **InsertSymbol**。

**方法**

|                                                              | 名称                  | 说明                                                         |
| ------------------------------------------------------------ | --------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ChangeCase**        | 更改指定文本的大小写。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Characters**        | 返回一个 **TextRange** 对象，该对象代表指定的文本字符子集。有关计算或浏览一段文本中的字符的详细信息，请参阅 **TextRange** 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Copy**              | 将指定对象复制到剪贴板。                                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Cut**               | 删除指定对象并将其放到剪贴板中。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**            | 删除指定的对象。                                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Find**              | 在一个文本范围内查找指定的文本，并返回 [TextRange ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/TextRange/TextRange%20.htm#jsObject_TextRange)对象，该对象代表在其中找到指定的文本的第一个文本范围。如果找不到指定的文本，则返回 **Nothing**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **InsertAfter**       | 在指定文本范围的末尾添加一个字符串。返回一个代表添加文本的 **TextRange** 对象。省略参数时，该方法返回一个在指定范围末尾的长度为零的字符串。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **InsertBefore**      | 在指定文本范围的开头添加一个字符串。返回一个代表添加文本的 **TextRange** 对象。省略参数时，该方法返回一个在指定范围末尾的长度为零的字符串。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **InsertDateTime**    | 在指定的文本范围中插入日期和时间。返回一个代表所插入文本的 **TextRange** 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **InsertSlideNumber** | 将当前幻灯片的幻灯片编号插入指定的文本范围。返回一个代表幻灯片编号的 **TextRange** 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **InsertSymbol**      | 返回一个 [TextRange ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/TextRange/TextRange%20.htm#jsObject_TextRange)对象，该对象代表一个插入到指定文本范围的符号。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Lines**             | 返回一个 [TextRange ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/TextRange/TextRange%20.htm#jsObject_TextRange)对象，该对象代表指定的文本行子集。有关计算或浏览文本范围中文本行的信息，请参阅 **TextRange**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/TextRange/TextRange%20.htm#jsObject_TextRange)对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Paragraphs**        | 返回一个 **TextRange** 对象，该对象代表文本段落的指定子集。有关统计或遍历文本范围内的段落的信息，请参阅 [TextRange](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/TextRange/TextRange%20.htm#jsObject_TextRange)对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Paste**             | 将剪贴板上的文本粘贴到指定的文本范围中，并返回一个代表粘贴文本的 **TextRange** 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Replace**           | 在文本范围内查找特定文本，用指定的字符串替换查找到的文本，返回代表查找文本的第一处匹配内容的 **TextRange** 对象。如果找不到匹配的内容，则返回 **Nothing**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Select**            | 选择指定的对象。                                             |

**属性**

|                                                              | 名称                | 说明                                                         |
| ------------------------------------------------------------ | ------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**     | 返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示指定对象的创建者。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BoundHeight**     | 以磅为单位返回指定文本框的文本边界框的高度。只读 **Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BoundLeft**       | 以磅为单位返回指定文本框的文本边界框的左边界到幻灯片的左边界的距离。只读。**Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BoundTop**        | 以磅为单位返回指定文本框的文本边界框的上边界到幻灯片的上边界的距离。只读。**Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BoundWidth**      | 以磅为单位返回指定文本框的文本边界框的宽度。只读 **Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**           | 返回指定集合中的对象数目。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Font**            | 返回一个代表字符格式的 [Font ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Font/Font%20.htm#jsObject_Font)对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ParagraphFormat** | 返回一个**ParagraphFormat**对象，该对象代表指定文本的段落格式。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Text**            | 返回或设置一个 **String** 类型值，该值代表指定对象中包含的文本。可读/写。 |

**成员方法**

#### **TextRange.ChangeCase**

更改指定文本的大小写。

**语法**

**express.ChangeCase(Type)**

*express*   一个代表 **TextRange** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**     | **说明**               |
| -------- | ------------- | ---------------- | ---------------------- |
| *Type*   | 必选          | **PpChangeCase** | 指定更改大小写的方法。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `本示例将当前演示文稿的第一张幻灯片的标题设为大写。 Application.ActivePresentation.Slides.Item(1).Shapes.Title.TextFrame.TextRange.ChangeCase(ppCaseTitle)` |

#### **TextRange.Characters**

返回一个 **TextRange** 对象，该对象代表指定的文本字符子集。有关计算或浏览一段文本中的字符的详细信息，请参阅 **TextRange** 对象。

**语法**

**express.Characters(Start, Length)**

*express*   一个代表 **TextRange** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                   |
| -------- | ------------- | ------------ | -------------------------- |
| *Start*  | 可选          | **Long**     | 返回的范围中的第一个字符。 |
| *Length* | 可选          | **Long**     | 要返回的字符数。           |

**说明**

如果 **Start** 和 **Length** 都被忽略，则返回的范围从指定范围中的第一个字符开始，到最后一段结束。

如果指定 **Start** 但忽略 **Length**，则返回的范围包含一个字符。

如果指定 **Length** 但忽略 **Start**，则返回的范围从指定范围中的第一个字符开始。

如果 **Start** 大于指定文本中的字符数，则返回的范围从指定范围中的最后一个字符开始。

如果 **Length** 大于从文本的指定开始字符到末尾的字符数，则返回的范围包含所有这些字符。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置当前演示文稿中第一张幻灯片的第二个形状的文本，并将第二个字符设为偏移 20% 的下标。*/ function test() {   let myShape = Application.ActivePresentation.Slides.Item(1).Shapes.Item(2)   let charRange = myShape.TextFrame.TextRange.InsertBefore("H2O")   charRange.Characters(2).Font.BaselineOffset = -0.2 }  /*本示例会将第一张幻灯片上第二个形状的所有下标字符设为加粗。*/ function test() {   let Trng = Application.ActivePresentation.Slides.Item(1).Shapes.Item(2).TextFrame2.TextRange   for(let i=1;i <= Trng.Characters().Count;i++) {       let ft = Trng.Characters(i).Font       if(ft.Subscript) {           ft.Bold = true       }   } }` |

#### **TextRange.Copy**

将指定对象复制到剪贴板。

**语法**

**express.Copy()**

*express*   一个代表 **TextRange** 对象的变量。

**说明**

使用 **Paste** 方法粘贴剪贴板中的内容。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在第一个窗口中拷贝选定幻灯片到剪贴板然后将其粘贴到第二窗口的视图中。如果剪贴板的内容不能粘贴到第二个窗口的视图中（例如，如果试图将一个形状粘贴到幻灯片浏览视图中），则本示例失败。*/ function test() {   Windows.Item(1).Selection.Copy()   Windows.Item(2).View.Paste() }  /*本示例将当前演示文稿中第一张幻灯片的第一个和第二个形状复制到剪贴板，然后将该副本粘贴到第二张灯片。*/ function test() {   ActivePresentation.Slides.Item(1).Shapes.Range([1, 2]).Copy()   ActivePresentation.Slides.Item(2).Shapes.Paste() }  /*本示例将当前演示文稿的第一张幻灯片复制到剪贴板。*/ ActivePresentation.Slides.Item(1).Copy()  /*本示例将当前演示文稿中第一张幻灯片的第一个形状中的文本复制到剪贴板。*/ ActivePresentation.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.Copy()` |

#### **TextRange.Cut**

删除指定对象并将其放到剪贴板中。

**语法**

**express.Cut()**

*express*   一个代表 **TextRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例删除第一个窗口中选中的对象并将它的一个副本放到剪贴板中。*/ Windows.Item(1).Selection.Cut()  /*本示例从当前演示文稿的第一张幻灯片中删除第一个和第二个形状，并将它们的副本放到剪贴板，然后将该副本粘贴到第二张幻灯片。*/ function test() {   ActivePresentation.Slides.Item(1).Shapes.Range([1, 2]).Cut()   ActivePresentation.Slides.Item(2).Shapes.Paste() }  /*本示例从当前演示文稿中删除第一张幻灯片并将它的一个副本放到剪贴板中。*/ ActivePresentation.Slides.Item(1).Cut()  /*本示例删除当前演示文稿中第一张幻灯片的第一个形状中的文本，并将它的一个副本放到剪贴板*/ ActivePresentation.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.Cut()` |

#### **TextRange.Delete**

删除指定的对象。

**语法**

**express.Delete()**

*express*   一个代表 **TextRange** 对象的变量。

**说明**

如果试图删除表格中唯一一个现有行或列，将导致运行时错误。

#### **TextRange.Find**

在一个文本范围内查找指定的文本，并返回 [TextRange ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/TextRange/TextRange%20.htm#jsObject_TextRange)对象，该对象代表在其中找到指定的文本的第一个文本范围。如果找不到指定的文本，则返回 **Nothing**。

**语法**

**express.Find(FindWhat, After, MatchCase, WholeWords)**

*express*   一个代表 **TextRange** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型**    | **说明**                                                     |
| ------------ | ------------- | --------------- | ------------------------------------------------------------ |
| *FindWhat*   | 必选          | **String**      | 要搜索的文本。                                               |
| *After*      | 可选          | **Long**        | 指定文本范围内的字符位置，将从该字符开始搜索下一处 FindWhat 匹配内容。例如，如果要从文本范围的第五个字符开始搜索，可指定 After 为 4。如果忽略该参数，则将文本范围的第一个字符作为搜索的起始点。 |
| *MatchCase*  | 可选          | **MsoTriState** | 如果此属性为 MsoTrue，则搜索时区分字符的大小写。             |
| *WholeWords* | 可选          | **MsoTriState** | 如果此属性为 MsoTrue，则搜索时仅查找整个词，而不搜索较长单词的部分字符。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `本示例在当前演示文稿中查找所有“CompanyX”字符串，并将其格式设为加粗。 function test() {   for(let i=1;i <= Application.ActivePresentation.Slides.Count;i++) {       let sli = Application.ActivePresentation.Slides       for(let j=1;j <= sli.Item(i).Shapes.Count;j++) {           if(sli.Item(i).Shapes.Item(j).HasTextFrame) {               let txtRng = sli.Item(i).Shapes.Item(j).TextFrame.TextRange               let foundText = txtRng.Find("CompanyX")               while(foundText != null) {                   foundText.Font.Bold = true                   let foundText = txtRng.Find("CompanyX", foundText.Start + foundText.Length-1)               }           }       }   } }` |

#### **TextRange.InsertAfter**

在指定文本范围的末尾添加一个字符串。返回一个代表添加文本的 **TextRange** 对象。省略参数时，该方法返回一个在指定范围末尾的长度为零的字符串。

**语法**

**express.InsertAfter(NewText)**

*express*   一个代表 **TextRange** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型** | **说明**                             |
| --------- | ------------- | ------------ | ------------------------------------ |
| *NewText* | 可选          | **String**   | 要插入的文字。默认值为一个空字符串。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将字符串“: Test version”追加到活动演示文稿第一张幻灯片标题的末尾。*/ function test() {   let myShape = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1)   myShape.TextFrame.TextRange.InsertAfter(": Test version") }  /*本示例将剪贴板的内容追加到第一张幻灯片标题的末尾。*/ Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.InsertAfter().Paste() ` |

#### **TextRange.InsertBefore**

在指定文本范围的开头添加一个字符串。返回一个代表添加文本的 **TextRange** 对象。省略参数时，该方法返回一个在指定范围末尾的长度为零的字符串。

**语法**

**express.InsertBefore(NewText)**

*express*   一个代表 **TextRange** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型** | **说明**                             |
| --------- | ------------- | ------------ | ------------------------------------ |
| *NewText* | 可选          | **String**   | 要添加的文本。默认值为一个空字符串。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将字符串“Test version:”添加到活动演示文稿第一张幻灯片标题的开头。*/ function test() {   let myShape = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1)   myShape.TextFrame.TextRange.InsertBefore("Test version: ") }  /*本示例将剪贴板中的内容添加到活动演示文稿第一张幻灯片标题的开头。*/ Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.InsertBefore().Paste()` |

#### **TextRange.InsertDateTime**

在指定的文本范围中插入日期和时间。返回一个代表所插入文本的 **TextRange** 对象。

**语法**

**express.InsertDateTime(DateTimeFormat, InsertAsField )**

*express*   一个代表 **TextRange** 对象的变量。

**参数**

| **名称**         | **必选/可选** | **数据类型**         | **说明**                                           |
| ---------------- | ------------- | -------------------- | -------------------------------------------------- |
| *DateTimeFormat* | 必选          | **PpDateTimeFormat** | 日期和时间的格式。                                 |
| *InsertAsField*  | 可选          | **MsoTriState**      | 决定是否每次打开演示文稿时都更新插入的日期和时间。 |

**说明**

| PpDateTimeFormat 可以是下列 PpDateTimeFormat 常量之一。 |
| ------------------------------------------------------- |
| **ppDateTimeddddMMMMddyyyy**                            |
| **ppDateTimedMMMMyyyy**                                 |
| **ppDateTimedMMMyy**                                    |
| **ppDateTimeFormatMixed**                               |
| **ppDateTimeHmm**                                       |
| **ppDateTimehmmAMPM**                                   |
| **ppDateTimeHmmss**                                     |
| **ppDateTimehmmssAMPM**                                 |
| **ppDateTimeMdyy**                                      |
| **ppDateTimeMMddyyHmm**                                 |
| **ppDateTimeMMddyyhmmAMPM**                             |
| **ppDateTimeMMMMdyyyy**                                 |
| **ppDateTimeMMMMyy**                                    |
| **ppDateTimeMMyy**                                      |

| MsoTriState 可以是下列 MsoTriState 常量之一。          |
| ------------------------------------------------------ |
| **msoCTrue**                                           |
| **msoFalse** 默认值。                                  |
| **msoTriStateMixed**                                   |
| **msoTriStateToggle**                                  |
| **msoTrue** 每次打开演示文稿时，更新插入的日期和时间。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `本示例将日期和时间插入在活动演示文稿第一张幻灯片第二个形状的第一段第一句之后。 function test() {   let sh = Application.ActivePresentation.Slides.Item(1).Shapes.Item(2)   let sentOne = sh.TextFrame.TextRange.Paragraphs(1).Sentences(1)   sentOne.InsertAfter().InsertDateTime(ppDateTimeMdyy) }` |

#### **TextRange.InsertSlideNumber**

将当前幻灯片的幻灯片编号插入指定的文本范围。返回一个代表幻灯片编号的 **TextRange** 对象。

**语法**

**express.InsertSlideNumber()**

*express*   一个代表 **TextRange** 对象的变量。

**说明**

当前幻灯片的幻灯片编号改变时，已插入的幻灯片编号将自动更新。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `本示例将当前幻灯片的编号插入到活动演示文稿第一张幻灯片第二个形状的第一段第一句之后。 function test() {   let sh = Application.ActivePresentation.Slides.Item(1).Shapes.Item(2)   let sentOne = sh.TextFrame.TextRange.Paragraphs(1).Sentences(1)   sentOne.InsertAfter().InsertSlideNumber() }` |

#### **TextRange.InsertSymbol**

返回一个 [TextRange ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/TextRange/TextRange%20.htm#jsObject_TextRange)对象，该对象代表一个插入到指定文本范围的符号。

**语法**

**express.InsertSymbol(FontName, CharNumber, Unicode)**

*express*   一个代表 **TextRange** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型**    | **说明**                                               |
| ------------ | ------------- | --------------- | ------------------------------------------------------ |
| *FontName*   | 必选          | **String**      | 字体名称。                                             |
| *CharNumber* | 必选          | **Long**        | Unicode 或 ASCII 字符编号。                            |
| *Unicode*    | 必选          | **MsoTriState** | 指定 CharNumber 参数代表 ASCII 字符还是 Unicode 字符。 |

**说明**

| **MsoTriState** 可以是下列 **MsoTriState** 常量之一。        |
| ------------------------------------------------------------ |
| **msoCTrue** 不适用于此方法。                                |
| **msoFalse** 默认值。**CharNumber** 参数代表一个 ASCII 字符编号。 |
| **msoTriStateMixed** 不适用于此方法。                        |
| **msoTriStateToggle** 不适用于此方法。                       |
| **msoTrue****CharNumber** 参数代表一个 Unicode 字符。        |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `本示例在活动演示文稿第一张幻灯片上的新文本框的第一段第一句后插入注册商标符号。 function test() {     //Add text box     let txtBox = Application.ActivePresentation.Slides.Item(1).Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 100, 100)     //Add symbol to text box     txtBox.TextFrame.TextRange.InsertSymbol("Symbol", 226) }` |

#### **TextRange.Lines**

返回一个 [TextRange ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/TextRange/TextRange%20.htm#jsObject_TextRange)对象，该对象代表指定的文本行子集。有关计算或浏览文本范围中文本行的信息，请参阅 **TextRange**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/TextRange/TextRange%20.htm#jsObject_TextRange)对象。

**语法**

**express.Lines(Start, Length)**

*express*   一个代表 **TextRange** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**               |
| -------- | ------------- | ------------ | ---------------------- |
| *Start*  | 必选          | **Long**     | 返回的范围中的第一行。 |
| *Length* | 必选          | **Long**     | 要返回的行数。         |

**说明**

如果 **Start** 和 **Length** 都被忽略，则返回的范围从指定范围中的第一行开始，到最后一段结束。

如果指定 **Start** 但忽略 **Length**，则返回的范围包含一行文本。

如果指定 **Length** 但忽略 **Start**，则返回的范围从指定范围中的第一行开始。

如果 **Start** 大于指定文本中的行数，则返回的范围从指定范围中的最后一行开始。

如果 **Length** 大于从文本的指定起始行到末尾的行数，则返回的范围包含所有这些行。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `本示例将当前演文示稿第一张幻灯片第二个形状中第二段的前两行设为斜体。 Application.ActivePresentation.Slides.Item(1).Shapes.Item(2).TextFrame.TextRange.Paragraphs(2).Lines(1, 2).Font.Italic = true ` |

#### **TextRange.Paragraphs**

返回一个 **TextRange** 对象，该对象代表文本段落的指定子集。有关统计或遍历文本范围内的段落的信息，请参阅 [TextRange](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/TextRange/TextRange%20.htm#jsObject_TextRange)对象。

**语法**

**express.Paragraphs(Start, Length)**

*express*   一个代表 **TextRange** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**               |
| -------- | ------------- | ------------ | ---------------------- |
| *Start*  | 可选          | **Long**     | 返回的范围中的第一段。 |
| *Length* | 可选          | **Long**     | 要返回的段落数。       |

**说明**

有关统计或遍历文本范围内的段落的信息，请参阅 [TextRange ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/TextRange/TextRange%20.htm#jsObject_TextRange)对象。

如果 *Start* 和 *Length* 都省略，则返回的范围从第一段开始，到指定范围的最后一段结束。

如果指定 *Start* 但忽略 *Length*，则返回的范围包含一个段落。

如果指定 *Length* 但忽略 *Start*，则返回的范围从指定范围中的第一段开始。

如果 *Start* 参数大于指定文本的段落数，则返回的范围从指定范围中的最后一段开始。

如果 *Length* 大于从指定起始段落到文字结尾的段落数，则返回的范围包含所有这些段落。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `以下示例将当前演文示稿第一张幻灯片第二个形状的第二段前两行设为倾斜。 Application.ActivePresentation.Slides.Item(1).Shapes.Item(2).TextFrame.TextRange.Paragraphs(2).Lines(1, 2).Font.Italic = true` |

#### **TextRange.Paste**

将剪贴板上的文本粘贴到指定的文本范围中，并返回一个代表粘贴文本的 **TextRange** 对象。

**语法**

**express.Paste()**

*express*   一个代表 **TextRange** 对象的变量。

**说明**

将剪贴板内容粘贴到窗口中之前，请使用 **ViewType** 属性设置窗口视图。下表显示了可以粘贴到每个视图中的内容。

| 视图                   | 可以从剪贴板中粘贴以下内容                                   |
| ---------------------- | ------------------------------------------------------------ |
| 幻灯片视图或备注页视图 | 形状、文本或整张幻灯片。如果从剪贴板粘贴一个幻灯片，该幻灯片的图像将作为嵌入对象被插入到幻灯片、母版或备注页中；如果选中形状，粘贴的文本将附加到形状文本之后；如果选中文本，粘贴的文本将替换选中的文本；如果选中任何其他对象，粘贴的文本将被放到它自己的文本框中。粘贴的形状将被放到 z 顺序的最上端且不会替换选中的形状。 |
| 大纲视图               | 文本或整张幻灯片。不能向大纲视图粘贴形状。粘贴的幻灯片将被插到光标所在的幻灯片之前。 |
| 幻灯片浏览视图         | 整张幻灯片。不能向幻灯片浏览视图粘贴形状或文本。粘贴的幻灯片将被插到光标处或演示文稿中最后选中的一张幻灯片之后。 |

 

#### **TextRange.Replace**

在文本范围内查找特定文本，用指定的字符串替换查找到的文本，返回代表查找文本的第一处匹配内容的 **TextRange** 对象。如果找不到匹配的内容，则返回 **Nothing**。

**语法**

**express.Replace(FindWhat, ReplaceWhat, After, MatchCase, WholeWords)**

*express*   一个代表 **TextRange** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型**    | **说明**                                                     |
| ------------- | ------------- | --------------- | ------------------------------------------------------------ |
| *FindWhat*    | 必选          | **String**      | 要搜索的文本。                                               |
| *ReplaceWhat* | 必选          | **String**      | 用来替换查找到的文本的文本。                                 |
| *After*       | 可选          | **Integer**     | 指定文本范围内开始搜索下一处 FindWhat 匹配内容的字符位置。例如，如果要从文本范围的第五个字符开始搜索，则可指定 After 为 4。如果省略此参数，则将文本范围的第一个字符作为搜索的起点。 |
| *MatchCase*   | 可选          | **MsoTriState** | 确定是否区分大小写。                                         |
| *WholeWords*  | 可选          | **MsoTriState** | 确定是否只查找全字匹配的内容。                               |

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| --------------------------------------------- |
| **msoCTrue**                                  |
| **msoFalse** 默认值。                         |
| **msoTriStateMixed**                          |
| **msoTriStateToggle**                         |
| **msoTrue** 区分字符的大小写。                |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `本示例在活动文档内的所有形状中，查找每一处全字匹配“like”的内容，并将其替换为“NOT LIKE”。 function test() {     let oSld = Application.ActivePresentation.Slides.Item(1)        for(let i=1;i <= oSld.Shapes.Count;i++) {         let oTxtRng = oSld.Shapes.Item(i).TextFrame.TextRange         let oTmpRng = oTxtRng.Replace("like","NOT LIKE",undefined,undefined,true)         while(oTmpRng != null) {             let oTxtRng = oTxtRng.Characters(oTmpRng.Start + oTmpRng.Length, oTxtRng.Length)             let oTmpRng = oTxtRng.Replace("like","NOT LIKE",undefined,undefined,true)         }     } }` |

#### **TextRange.Select**

选择指定的对象。

**语法**

**express.Select()**

*express*   一个代表 **TextRange** 对象的变量。

**说明**

如果做出不适用于当前视图的选择，此代码将失败。例如，可以在幻灯片浏览视图中而不能在幻灯片视图中选择幻灯片。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例选择活动演示文稿中第一张幻灯片标题的前五个字符。*/ ActivePresentation.Slides.Item(1).Shapes.Title.TextFrame.TextRange.Characters(1, 5).Select() /*本示例选择活动演示文稿中的第一张幻灯片。*/ ActivePresentation.Slides.Item(1).Select() /*本示例选取一个已添加到演示文稿的新幻灯片中的表格。该表格具有三行和三列。*/ function test() {   let sli = Presentations.Add().Slides   sli.Add(1, ppLayoutBlank).Shapes.AddTable(3, 3).Select() }` |

**成员属性**

#### **TextRange.Application**

返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **TextRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。*/ function test(pptPres) {     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") }  /*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。*/ function test() {   let shpOle = ActivePresentation.Slides.Item(1).Shapes   for(let i = 1; i <= shpOle.Count; i++) {       if(shpOle.Item(i).Type == msoLinkedOLEObject) {           MsgBox(shpOle.Item(i).OLEFormat.Application.Name)       }   } }` |

#### **TextRange.BoundHeight**

以磅为单位返回指定文本框的文本边界框的高度。只读 **Single** 类型。

**语法**

**express.BoundHeight**

*express*   一个代表 **TextRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `本示例在活动演示文稿的第一张幻灯片中添加一个圆角矩形，该矩形与第一个形状的文本边界框的大小相同。 function test() {   let shape = Application.ActivePresentation.Slides.Item(1).Shapes   let tr = shape.Item(1).TextFrame.TextRange   let roundRect = shape.AddShape(msoShapeRoundedRectangle, tr.BoundLeft, tr.BoundTop, tr.BoundWidth, tr.BoundHeight)   let fill = roundRect.Fill   fill.ForeColor.RGB = 255, 0, 128   fill.Transparency = 0.75 }` |

#### **TextRange.BoundLeft**

以磅为单位返回指定文本框的文本边界框的左边界到幻灯片的左边界的距离。只读。**Single** 类型。

**语法**

**express.BoundLeft**

*express*   一个代表 **TextRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `本示例在活动演示文稿的第一张幻灯片中添加一个圆角矩形，该矩形与第一个形状的文本边界框的大小相同。 function test() {   let shape = Application.ActivePresentation.Slides.Item(1).Shapes   let tr = shape.Item(1).TextFrame.TextRange   let roundRect = shape.AddShape(msoShapeRoundedRectangle, tr.BoundLeft, tr.BoundTop, tr.BoundWidth, tr.BoundHeight)   let fill = roundRect.Fill   fill.ForeColor.RGB = 255, 0, 128   fill.Transparency = 0.75 }` |

#### **TextRange.BoundTop**

以磅为单位返回指定文本框的文本边界框的上边界到幻灯片的上边界的距离。只读。**Single** 类型。

**语法**

**express.BoundTop**

*express*   一个代表 **TextRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `本示例在活动演示文稿的第一张幻灯片中添加一个圆角矩形，该矩形与第一个形状的文本边界框的大小相同。 function test() {   let shape = Application.ActivePresentation.Slides.Item(1).Shapes   let tr = shape.Item(1).TextFrame.TextRange   let roundRect = shape.AddShape(msoShapeRoundedRectangle, tr.BoundLeft, tr.BoundTop, tr.BoundWidth, tr.BoundHeight)   let fill = roundRect.Fill   fill.ForeColor.RGB = 255, 0, 128   fill.Transparency = 0.75 }` |

#### **TextRange.BoundWidth**

以磅为单位返回指定文本框的文本边界框的宽度。只读 **Single** 类型。

**语法**

**express.BoundWidth**

*express*   一个代表 **TextRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `本示例在活动演示文稿的第一张幻灯片中添加一个圆角矩形，该矩形与第一个形状的文本边界框的大小相同。 function test() {   let shape = Application.ActivePresentation.Slides.Item(1).Shapes   let tr = shape.Item(1).TextFrame.TextRange   let roundRect = shape.AddShape(msoShapeRoundedRectangle, tr.BoundLeft, tr.BoundTop, tr.BoundWidth, tr.BoundHeight)   let fill = roundRect.Fill   fill.ForeColor.RGB = 255, 0, 128   fill.Transparency = 0.75 }` |

#### **TextRange.Count**

返回指定集合中的对象数目。

**语法**

**express.Count**

*express*   一个代表 **TextRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `以下示例关闭除当前窗口外的所有窗口。 function test() {   let win = Application.Windows   for (let i = 2; i <= win.Count; i++) {       win.Item(2).Close()   } }` |

#### **TextRange.Font**

返回一个代表字符格式的 [Font ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Font/Font%20.htm#jsObject_Font)对象。只读。

**语法**

**express.Font**

*express*   一个代表 **TextRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例为当前演示文稿第一张幻灯片的第一个形状中的文本设置格式。*/ function test() {   let shape = ActivePresentation.Slides.Item(1).Shapes.Item(1)   let font = shape.TextFrame.TextRange.Font   font.Size = 48   font.Name = "Palatino"   font.Bold = true   font.Color.RGB = 55, 127, 255 } /*以下示例设置第一张幻灯片第二个形状中项目符号的颜色和字体名称。*/ function test() {   let shape = ActivePresentation.Slides.Item(1).Shapes.Item(2)   let bullet = shape.TextFrame.TextRange.ParagraphFormat.Bullet   bullet.Visible = true   let font = bullet.Font   font.Name = "Palatino"   font.Color.RGB = 0, 0, 255 }` |

#### **TextRange.ParagraphFormat**

返回一个**ParagraphFormat**对象，该对象代表指定文本的段落格式。只读。

**语法**

**express.ParagraphFormat**

*express*   一个代表 **TextRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `以下示例设置当前演示文稿第一张幻灯片第二个形状中每段的段前、段内和段后行间距。 function test() {   let shape = Application.ActivePresentation.Slides.Item(2).Shapes.Item(2)   let format = shape.TextFrame.TextRange.ParagraphFormat   format.LineRuleWithin = msoTrue   format.SpaceWithin = 1.4   format.LineRuleBefore = msoTrue   format.SpaceBefore = 0.25   format.LineRuleAfter = msoTrue   format.SpaceAfter = 0.75 }` |

#### **TextRange.Text**

返回或设置一个 **String** 类型值，该值代表指定对象中包含的文本。可读/写。

**语法**

**express.Text**

*express*   一个代表 **TextRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `本示例设置当前演示文稿第一张幻灯片中标题的文本和字形。 function test() {   let myPres = Application.ActivePresentation   let range = myPres.Slides.Item(1).Shapes.Title.TextFrame.TextRange   range.Text = "Welcome!"   range.Font.Italic = true }` |

适用环境：web

适用平台：windows/linux