**Border**



代表一个对象的一个边框。**Border** 对象是 **Borders** 集合的成员。

**说明**

使用 **Borders**(*index*) 来返回单个 **Border** 对象（其中 *index* 标识边框）。Index 可以是某个 **WdBorderType** 常量。由于选择或安装的语言支持（如美国英语）不同，有些 **WdBorderType** 常量可能无法使用。

使用[**LineStyle**](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Border/Border%20.htm#Border.LineStyle)属性将边框线应用于 **Border** 对象。以下示例在活动文档的第一段下添加双线边框。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { 	let borders = Application.ActiveDocument.Paragraphs.Item(1).Borders.Item(wdBorderBottom) 	borders.LineStyle = wdLineStyleDouble 	borders.LineWidth = wdLineWidth025pt }` |

以下示例将单线边框加在选定内容的首字符四周。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { 	let characters = Application.Selection.Characters.Item(1) 	characters.Font.Size = 36 	characters.Borders.Enable = true  }` |

以下示例在第一节中每一页的四周添加艺术型边框。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { 	let borders = Application.ActiveDocument.Sections.Item(1).Borders 	for(let i=1;i<=borders.Count;++i){ 	    borders.Item(i).ArtStyle = wdArtSeattle 	    borders.Item(i).ArtWidth = 20 	}  }` |

不能将 **Border** 对象添加到 **Borders** 集合。 **Borders** 集合中的成员数是有限的，具体数目因对象类型而异。例如，表格在 **Borders** 集合中有六个元素，而段落有四个。

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个 **Application** 对象，该对象代表 WPS 应用程序。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ArtStyle**    | 返回或设置文档的艺术型页面边框。**WdPageBorderArt** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ArtWidth**    | 返回或设置指定艺术型边框的宽度（以磅为单位）。**Number** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Color**       | 返回或设置指定 **Border** 对象的 24 位颜色。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ColorIndex**  | 返回或设置指定边框或字体对象的颜色。**WdColorIndex** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。**Long** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Inside**      | 如果指定对象可以使用内部边框，则该属性值为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LineStyle**   | 返回或设置指定对象的边框样式。**WdLineStyle** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LineWidth**   | 返回或设置对象边框的线条宽度。可读写。                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回一个 **Object** 类型的值，该值代表指定 **Border** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Visible**     | 如果指定对象可见，则该属性值为 **True**。**Boolean** 类型，可读写。 |

**成员属性**

#### **Border.Application**

返回一个 **Application** 对象，该对象代表 WPS 应用程序。

**语法**

**express.Application**

*express*   一个代表 **Border** 对象的变量。

#### **Border.ArtStyle**

返回或设置文档的艺术型页面边框。**WdPageBorderArt** 类型，可读写。

**语法**

**express.ArtStyle**

*express*   一个代表 **Border** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例为选定内容的第一节的每个页面添加由黑点构成的边框 */ function test() { 	for(let i=1;i <= Application.Selection.Sections.Item(1).Borders.Count;i++) { 	    Application.Selection.Sections.Item(1).Borders.Item(i).ArtStyle = wdArtBasicBlackDots 	    Application.Selection.Sections.Item(1).Borders.Item(i).ArtWidth = 6 	} }  /* 本示例为活动文档中的第一节的每个页面添加由特定图片所构成的边框 */ function test() { 	let Sec = Application.ActiveDocument.Sections.Item(1) 	Sec.Borders.AlwaysInFront = true 	for(let i=1;i <= Application.ActiveDocument.Sections.Item(1).Borders.Count;i++) { 	    Application.ActiveDocument.Sections.Item(1).Borders.Item(i).ArtStyle = wdArtPeople 	    Application.ActiveDocument.Sections.Item(1).Borders.Item(i).ArtWidth = 15 	} }` |

#### **Border.ArtWidth**

返回或设置指定艺术型边框的宽度（以磅为单位）。**Number** 类型，可读写。

**语法**

**express.ArtWidth**

*express*   一个代表 **Border** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例为选定内容的第一节的每个页面添加由黑点构成的宽为 6 磅的边框 */ function test() { 	for(let i=1;i <= Application.Selection.Sections.Item(1).Borders.Count;i++) { 	    let Bor = Application.Selection.Sections.Item(1).Borders.Item(i) 	    Bor.ArtStyle = wdArtBasicBlackDots 	    Bor.ArtWidth = 6 	} }` |

#### **Border.Color**

返回或设置指定 **Border** 对象的 24 位颜色。

**语法**

**express.Color**

*express*   一个代表 **Border** 对象的变量。

**说明**

可以是任何有效的 **WdColor** 常量

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例在第一个表格中每个单元格周围添加靛蓝色的点划线边框 */ function test() { 	if(ActiveDocument.Tables.Count >= 1) { 	    for(let i=1;i <= ActiveDocument.Tables.Item(1).Borders.Count;i++) { 	        ActiveDocument.Tables.Item(1).Borders.Item(i).Color = wdColorIndigo 	        ActiveDocument.Tables.Item(1).Borders.Item(i).LineStyle = wdLineStyleDashDot 	        ActiveDocument.Tables.Item(1).Borders.Item(i).LineWidth = wdLineWidth075pt 	    } 	} }` |

#### **Border.ColorIndex**

返回或设置指定边框或字体对象的颜色。**WdColorIndex** 类型，可读写。

**语法**

**express.ColorIndex**

*express*   一个代表 **Border** 对象的变量。

**说明**

**wdByAuthor** 常量对边框和字体对象无效。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例为第一个表格的每个单元格添加红虚线边框 */ function test() { 	if(Application.ActiveDocument.Tables.Count >= 1) { 	    for(let i=1;i <= Application.ActiveDocument.Tables.Item(1).Borders.Count;i++) { 	        let Bor = Application.ActiveDocument.Tables.Item(1).Borders.Item(i) 	        Bor.ColorIndex = wdRed 	        Bor.LineStyle = wdLineStyleDashDot 	        Bor.LineWidth = wdLineWidth075pt 	    } 	} }` |

#### **Border.Creator**

返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。**Long** 类型，只读。

**语法**

**express.Creator**

*express*   一个代表 **Border** 对象的变量。

**说明**

如果对象是在 WPS 中创建的，则 **Creator** 属性返回十六进制数 4D535744，代表字符串“WPS”。该属性主要是为 Macintosh 机的应用设计的，在该机上每个应用程序都有一个四字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参考 WPS OfficeMacintosh Edition 中的语言参考帮助。

| 注释                                    |
| --------------------------------------- |
| 该值也可用常量 **wdCreatorCode** 表示。 |

#### **Border.Inside**

如果指定对象可以使用内部边框，则该属性值为 **True**。**Boolean** 类型，只读。

**语法**

**express.Inside**

*express*   一个代表 **Border** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 如果当前选定内容支持内部边框（即选定了多个段落或单元格），则本示例为其设置单线型内部边框 */ function test() { 	for(let i=1;i <= Application.Selection.Borders.Count;i++) { 	    if(Application.Selection.Borders.Item(i).Inside == true) { 	        Application.Selection.Borders.Item(i).LineStyle = wdLineStyleSingle 	    } 	} }` |

#### **Border.LineStyle**

返回或设置指定对象的边框样式。**WdLineStyle** 类型，可读写。

**语法**

**express.LineStyle**

*express*   一个代表 **Border** 对象的变量。

**说明**

如果是为单个字符或单词区域设置 **LineStyle** 属性，应选用字符边框。

如果是为一个或几个段落区域设置 **LineStyle** 属性，应选用段落边框。连续段落之间的边框可应用 **InsideLineStyle** 属性。

如果是为节设置 **LineStyle** 属性，应选用节中的页面边框。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 如果所选内容是一个段落，或折叠起来的选定内容，本示例将在当前段落之前添加 0.75 磅宽的单线段落边框。如果所选内容中不包括段落，则为所选文字添加边框 */ function test() { 	let Bor = Application.Selection.Borders.Item(wdBorderTop) 	Bor.LineStyle = wdLineStyleSingle 	Bor.LineWidth = wdLineWidth075pt }` |

#### **Border.LineWidth**

返回或设置对象边框的线条宽度。可读写。

**语法**

**express.LineWidth**

*express*   一个代表 **Border** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例为活动文档的第一个表格的第一行添加下边框 */ function test() { 	if(Application.ActiveDocument.Tables.Count >= 1) { 	    let Bor = Application.ActiveDocument.Tables.Item(1).Rows.Item(1).Borders.Item(wdBorderBottom) 	    Bor.LineStyle = wdLineStyleDouble 	    Bor.LineWidth = wdLineWidth050pt 	} }` |

#### **Border.Parent**

返回一个 **Object** 类型的值，该值代表指定 **Border** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Border** 对象的变量。

#### **Border.Visible**

如果指定对象可见，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.Visible**

*express*   一个代表 **Border** 对象的变量。

**说明**

对于任何对象，如果其 **Visible** 属性值为 **False**，则某些方法和属性可能无法使用。

适用环境：web

适用平台：windows/linux