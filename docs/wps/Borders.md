**Border**



**Border** 对象的集合，这些对象代表对象的边框。

**方法**

|                                                              | 名称                              | 说明                                       |
| ------------------------------------------------------------ | --------------------------------- | ------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ApplyPageBordersToAllSections** | 将指定的页面边框格式应用于文档中的所有节。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**                          | 返回指定区域或选定内容中的一个边框。       |

**属性**

|                                                              | 名称                          | 说明                                                         |
| ------------------------------------------------------------ | ----------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AlwaysInFront**             | 如果在文档文本之前显示页面边框，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**               | 返回一个 **Application** 对象，该对象代表 WPS 应用程序。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**                     | 返回 **Borders** 集合中的项目数。**Long** 类型，只读。       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                   | 返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。**Long** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DistanceFrom**              | 此属性返回或设置数值，该值指明指定的页面边框是从页面边缘还是从环绕的文本开始度量。**WdBorderDistanceFrom**，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DistanceFromBottom**        | 返回或设置文本与下边框的距离（以磅为单位）。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DistanceFromLeft**          | 返回或设置文本与左边框的距离（以磅为单位）。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DistanceFromRight**         | 返回或设置文本的右边与右边框的距离（以磅为单位）。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DistanceFromTop**           | 返回或设置文本与上边框的距离（以磅为单位）。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Enable**                    | 返回或设置指定对象的边框格式。**Long** 类型，可读写。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EnableFirstPageInSection**  | 如果可为节内第一页设置页面边框，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EnableOtherPagesInSection** | 如果可为本节中除首页之外的所有页面设置页面边框，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HasHorizontal**             | 如果可将水平方向的边框应用于对象，则该属性值为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HasVertical**               | 如果可将垂直边框应用于指定对象，则该属性值为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **InsideColor**               | 返回或设置内部边框的 24 位颜色。可读写。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **InsideColorIndex**          | 返回或设置内部边框的颜色。**WdColorIndex** 类型，可读写。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **InsideLineStyle**           | 返回或设置指定对象的内部边框。                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **InsideLineWidth**           | 返回或设置对象内部边框的线条宽度。                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **JoinBorders**               | 如果删除段落和表格边缘的垂直边框，以使水平边框与页面边框相连，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OutsideColor**              | 返回或设置外部边框的 24 位颜色。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OutsideColorIndex**         | 返回或设置外部边框的颜色。**WdColorIndex** 类型，可读写。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OutsideLineStyle**          | 返回或设置指定对象的外部边框。                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OutsideLineWidth**          | 返回或设置对象外部边框的线条宽度。可读写。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                    | 返回一个 **Object** 类型的值，该值代表指定 **Borders** 集合中的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shadow**                    | 如果为 **True**，则指定的边框设置为阴影格式。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SurroundFooter**            | 如果页面边框环绕文档的页脚，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SurroundHeader**            | 如果页面边框环绕文档页眉，则该属性值为 **True**。**Boolean** 类型，可读写。 |

**成员方法**

#### **Borders.ApplyPageBordersToAllSections**

将指定的页面边框格式应用于文档中的所有节。

**语法**

**express.ApplyPageBordersToAllSections()**

*express*   一个代表 **Borders** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为活动文档的所有节添加单线型页面边框*/ function test() {     let sec = Application.ActiveDocument.Sections.Item(1)     for(let i=1;i<=sec.Borders.Count;i++){ 	sec.Borders.Item(i).LineStyle = wdLineStyleSingle 	sec.Borders.Item(i).LineWidth = wdLineWidth050pt     }     sec.Borders.ApplyPageBordersToAllSections() }` |

#### **Borders.Item**

返回指定区域或选定内容中的一个边框。

**语法**

**express.Item(Index)**

*express*   一个代表 **Borders** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**     | **说明**     |
| -------- | ------------- | ---------------- | ------------ |
| *Index*  | 必选          | **WdBorderType** | 返回的边框。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动文档的第一段之上插入双边框*/ function BorderItem() {     Application.ActiveDocument.Paragraphs.Item(1).Borders.Item(wdBorderTop).LineStyle = wdLineStyleDouble } ` |

**成员属性**

#### **Borders.AlwaysInFront**

如果在文档文本之前显示页面边框，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.AlwaysInFront**

*express*   一个代表 **Borders** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动文档中第一节的文档文本之前添加艺术型页面边框*/ function test() { 	let Sec = Application.ActiveDocument.Sections.Item(1) 	Sec.Borders.AlwaysInFront = true 	for(let i = 1; i <= Sec.Borders.Count; i++) { 	    let Bor =Sec.Borders.Item(i) 	    Bor.ArtStyle = wdArtPeople 	    Bor.ArtWidth = 15 	} }` |

#### **Borders.Application**

返回一个 **Application** 对象，该对象代表 WPS 应用程序。

**语法**

**express.Application**

*express*   一个代表 **Borders** 对象的变量。

#### **Borders.Count**

返回 **Borders** 集合中的项目数。**Long** 类型，只读。

**语法**

**express.Count**

*express*   一个代表 **Borders** 对象的变量。

#### **Borders.Creator**

返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。**Long** 类型，只读。

**语法**

**express.Creator**

*express*   一个代表 **Borders** 对象的变量。

**说明**

如果对象是在 WPS 中创建的，则 **Creator** 属性返回十六进制数 4D535744，代表字符串“WPS”。该属性主要是为 Macintosh 机的应用设计的，在该机上每个应用程序都有一个四字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参考 WPS OfficeMacintosh Edition 中的语言参考帮助。

#### **Borders.DistanceFrom**

此属性返回或设置数值，该值指明指定的页面边框是从页面边缘还是从环绕的文本开始度量。**WdBorderDistanceFrom**，可读写。

**语法**

**express.DistanceFrom**

*express*   一个代表 **Borders** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为活动文档中第一节的每个页面添加单线型边框，然后设置每个边框到页面边缘的距离*/ function test() { 	let Sec = Application.ActiveDocument.Sections.Item(1) 	for(let i = 1 ; i <= Sec.Borders.Count; i++) { 	    Sec.Borders.Item(i).LineStyle = wdLineStyleSingle 	    Sec.Borders.Item(i).LineWidth = wdLineWidth050pt 	} 	let Bor = Sec.Borders 	Bor.DistanceFrom = wdBorderDistanceFromPageEdge 	Bor.DistanceFromTop = 20 	Bor.DistanceFromLeft = 22 	Bor.DistanceFromBottom = 20 	Bor.DistanceFromRight = 22 }  /*本示例为所选内容的第一节的每个页面添加边框，然后将文本与页面边框的距离设置为 6 磅*/ function test() { 	let Sec = Application.Selection.Sections.Item(1) 	for(let i = 1; i <= Sec.Borders.Count; i++) { 	    Sec.Borders.Item(i).ArtStyle = wdArtSeattle 	    Sec.Borders.Item(i).ArtWidth = 22 	} 	let Bor = Sec.Borders 	Bor.DistanceFrom = wdBorderDistanceFromText 	Bor.DistanceFromTop = 6 	Bor.DistanceFromLeft = 6 	Bor.DistanceFromBottom = 6 	Bor.DistanceFromRight = 6 } ` |

#### **Borders.DistanceFromBottom**

返回或设置文本与下边框的距离（以磅为单位）。**Long** 类型，可读写。

**语法**

**express.DistanceFromBottom**

*express*   一个代表 **Borders** 对象的变量。

**说明**

将该属性与页边框结合使用，可以设置文本与页面下边框之间的距离，或页面下边缘与下边框之间的距离。这个距离是根据 **DistanceFrom** 属性的值测量的。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为活动文档的第一个段落添加边框，并将文本与下边框的距离设置为 6 磅。*/ function test() { 	let Bor = Application.ActiveDocument.Paragraphs.Item(1).Borders 	Bor.Enable = true 	Bor.DistanceFromBottom = 6 }  /*本示例为 Sales.doc 中的每张表格添加边框。然后将文本与上下边框的距离设置为 3 磅，将文本与左右边框的距离设置为 6 磅*/ function test() { 	for(let i = 1; i <= Application.Documents.Item("Sales.doc").Tables.Count; i++) { 	    let Bor = Application.Documents.Item("Sales.doc").Tables.Item(i).Borders 	    Bor.OutsideLineStyle = wdLineStyleSingle 	    Bor.OutsideLineWidth = wdLineWidth150pt 	    Bor.DistanceFromBottom = 3 	    Bor.DistanceFromTop = 3 	    Bor.DistanceFromLeft = 6 	    Bor.DistanceFromRight = 6 	} } ` |

#### **Borders.DistanceFromLeft**

返回或设置文本与左边框的距离（以磅为单位）。**Long** 类型，可读写。

**语法**

**express.DistanceFromLeft**

*express*   一个代表 **Borders** 对象的变量。

**说明**

将该属性与页边框结合使用，可设置文本与页面左边框之间的距离，或页面左边缘与页面左边框之间的距离。这个距离是根据 **DistanceFrom**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Borders/Borders%20.htm#Borders.DistanceFrom)属性的值测量的。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为活动文档中的每个图文框添加边框，并且将图文框与边框的距离设置为 5 磅。*/ function test() {         for(let i = 1; i <= Application.ActiveDocument.Frames.Count; i++) { 	    let Bor = Application.ActiveDocument.Frames.Item(i).Borders 	    Bor.Enable = true 	    Bor.DistanceFromLeft = 5 	    Bor.DistanceFromRight = 5 	    Bor.DistanceFromTop = 5 	    Bor.DistanceFromBottom = 5 	} } /*本示例为活动文档的第一个段落添加边框，并将文本与左边框的距离设置为 3 磅*/ function test() { 	let Bor = Application.ActiveDocument.Paragraphs.Item(1).Borders 	Bor.Enable = true 	Bor.DistanceFromLeft = 3 } ` |

#### **Borders.DistanceFromRight**

返回或设置文本的右边与右边框的距离（以磅为单位）。**Long** 类型，可读写。

**语法**

**express.DistanceFromRight**

*express*   一个代表 **Borders** 对象的变量。

**说明**

将该属性与页边框结合使用，可设置文本与右边框之间的距离，或页面右边缘与页面右边框之间的距离。这个距离是根据 **DistanceFrom**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Borders/Borders%20.htm#Borders.DistanceFrom)属性的值测量的。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为所选内容的每个段落添加边框，并将文本与右边框的距离设置为 3 磅。*/ function test() { 	let Bor = Application.Selection.Paragraphs.Borders 	Bor.Enable = true 	Bor.DistanceFromRight = 3 }  /*本示例为活动文档中第一节的每页添加单线型页面边框。然后将左右边框与相应的页面边缘的距离设置为 22 磅。*/ function test() { 	let Sec = Application.ActiveDocument.Sections.Item(1) 	for(let i = 1; i <= Sec.Borders.Count; i++) { 	    Sec.Borders.Item(i).LineStyle = wdLineStyleSingle 	    Sec.Borders.Item(i).LineWidth = wdLineWidth050pt 	} 	let Bor = Sec.Borders 	Bor.DistanceFrom = wdBorderDistanceFromPageEdge 	Bor.DistanceFromLeft = 22 	Bor.DistanceFromRight = 22 }` |

#### **Borders.DistanceFromTop**

返回或设置文本与上边框的距离（以磅为单位）。**Long** 类型，可读写。

**语法**

**express.DistanceFromTop**

*express*   一个代表 **Borders** 对象的变量。

**说明**

将该属性与页边框结合使用，可设置文本与上边框之间的距离，或页面上边缘与上边框之间的距离。这个距离是根据 **DistanceFrom**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Borders/Borders%20.htm#Borders.DistanceFrom)属性的值测量的。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为所选内容中的每个段落添加边框，并将文本与上边框的距离设置为 3 磅。*/ function test() { 	Application.Selection.Borders.Enable = true 	Application.Selection.Borders.DistanceFromTop = 3 }  /*本示例为所选内容中第一节的每页添加页面边框，并将文本与页面边框的距离设置为 6 磅。*/ function test() { 	let Sec = Application.Selection.Sections.Item(1) 	for(let i = 1; i <= Sec.Borders.Count; i++) { 	    Sec.Borders.Item(i).ArtStyle = wdArtSeattle 	    Sec.Borders.Item(i).ArtWidth = 22 	} 	let Bor = Sec.Borders 	Bor.DistanceFrom = wdBorderDistanceFromText 	Bor.DistanceFromTop = 6 	Bor.DistanceFromLeft = 6 	Bor.DistanceFromBottom = 6 	Bor.DistanceFromRight = 6 }` |

#### **Borders.Enable**

返回或设置指定对象的边框格式。**Long** 类型，可读写。

**语法**

**express.Enable**

*express*   一个代表 **Borders** 对象的变量。

**说明**

如果指定对象的全部或部分边框应用了边框格式，则 **Enable** 属性返回 **True** 或 **wdUndefined**。可将该属性值设置为 **True**、**False** 或 **WdLineStyle** 常量。

**Enable** 属性应用于指定对象的所有边框。如果为 **True**，则样式设置为默认样式，线条宽度设置为默认线条宽度。可以通过 **DefaultBorderLineWidth** 和 **DefaultBorderLineStyle** 属性来设置默认样式和默认线条宽度。

正如下例所示，将 **Enable** 属性的值设置为 **False**，即可删除对象中的所有边框。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActiveDocument.Tables.Item(1).Borders.Enable = false` |

如果要删除或应用单一边框，可先用 

Borders

(

index

) 返回单一边框，然后设置 

LineStyle

 属性（其中 

index

 为 

WdBorderType

 常量）。下例删除了 

```
Selection
```

 中的下边框。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Applicaiton.Selection.Borders.Item(wdBorderBottom).LineStyle = wdLineStyleNone` |

#### **Borders.EnableFirstPageInSection**

如果可为节内第一页设置页面边框，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.EnableFirstPageInSection**

*express*   一个代表 **Borders** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为选定内容的第一节的首页添加页面边框。*/ function test() { 	let Sec = Application.Selection.Sections.Item(1) 	Sec.Borders.EnableFirstPageInSection = true 	Sec.Borders.EnableOtherPagesInSection = false 	for(let i = 1; i <= Sec.Borders.Count; i++) {     	Sec.Borders.Item(i).ArtStyle = wdArtPeople     	Sec.Borders.Item(i).ArtWidth = 15 	}	 }` |

#### **Borders.EnableOtherPagesInSection**

如果可为本节中除首页之外的所有页面设置页面边框，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.EnableOtherPagesInSection**

*express*   一个代表 **Borders** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为选定内容的第一节除首页外的所有页面添加边框。*/ function test() { 	let Sec = Application.Selection.Sections.Item(1) 	Sec.Borders.EnableFirstPageInSection = false 	Sec.Borders.EnableOtherPagesInSection = true 	for(let i = 1; i <= Sec.Borders.Count; i++) { 	    Sec.Borders.Item(i).ArtStyle = wdArtBabyRattle 	    Sec.Borders.Item(i).ArtWidth = 22 	} } ` |

#### **Borders.HasHorizontal**

如果可将水平方向的边框应用于对象，则该属性值为 **True**。**Boolean** 类型，只读。

**语法**

**express.HasHorizontal**

*express*   一个代表 **Borders** 对象的变量。

**说明**

水平边框可以应用于包含两行或多行单元格的表格区域，或包含两个或多个段落的区域。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果选定内容支持水平方向的边框，则本示例为其设置单线型的水平边框。*/ function test() { 	if(Application.Selection.Borders.HasHorizontal == true) { 	    Application.Selection.Borders.Item(wdBorderHorizontal).LineStyle = wdLineStyleSingle 	} }` |

#### **Borders.HasVertical**

如果可将垂直边框应用于指定对象，则该属性值为 **True**。**Boolean** 类型，只读。

**语法**

**express.HasVertical**

*express*   一个代表 **Borders** 对象的变量。

**说明**

垂直边框可以应用于包含两列或多列单元格的表格区域。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果选定内容支持垂直边框，则本示例为其设置单线型的垂直边框。*/ function test() { 	if(Application.Selection.Borders.HasVertical == true) { 	    Application.Selection.Borders.Item(wdBorderVertical).LineStyle = wdLineStyleSingle 	} }` |

#### **Borders.InsideColor**

返回或设置内部边框的 24 位颜色。可读写。

**语法**

**express.InsideColor**

*express*   一个代表 **Borders** 对象的变量。

**说明**

该属性可以是任何有效的 **WdColor** 常量或 Visual Basic 的 **RGB** 函数返回的值。如果将 **InsideLineStyle** 属性设置为 **wdLineStyleNone** 或 **False**，则本属性的设置不产生任何效果。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动文档的前四个段落之间添加深红色边框。*/ function test() { 	let myDoc = Application.ActiveDocument 	let myRange = myDoc.Range(myDoc.Paragraphs.Item(1).Range.Start, myDoc.Paragraphs.Item(4).Range.End) 	let Bor = myRange.Borders 	Bor.InsideLineStyle = wdLineStyleSingle 	Bor.InsideLineWidth = wdLineWidth150pt 	Bor.InsideColor = wdDarkRed }` |

#### **Borders.InsideColorIndex**

返回或设置内部边框的颜色。**WdColorIndex** 类型，可读写。

**语法**

**express.InsideColorIndex**

*express*   一个代表 **Borders** 对象的变量。

**说明**

如果将 **InsideLineStyle** 属性设置为 **wdLineStyleNone** 或 **False**，则本属性的设置不产生任何效果。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动文档前四段之间添加红色边框。*/ function test() { 	let docActive = Application.ActiveDocument 	let rngTemp = docActive.Range(docActive.Paragraphs.Item(1).Range.Start, docActive.Paragraphs.Item(4).Range.End) 	let Bor = rngTemp.Borders 	Bor.InsideLineStyle = wdLineStyleSingle 	Bor.InsideLineWidth = wdLineWidth150pt 	Bor.InsideColorIndex = wdRed } ` |

#### **Borders.InsideLineStyle**

返回或设置指定对象的内部边框。

**语法**

**express.InsideLineStyle**

*express*   一个代表 **Borders** 对象的变量。

**说明**

如果指定对象应用了多种边框，则该属性返回 **wdUndefined**；否则，返回 **False** 或 **WdLineStyle** 常量。可将该属性值设置为 **True**、**False** 或 **WdLineStyle** 常量。

如果为 **True**，则样式设置为默认样式，线条宽度设置为默认线条宽度。可用 **DefaultBorderLineWidth** 和 **DefaultBorderLineStyle** 属性设置默认样式和线条宽度。

可用下列任何一种指令删除活动文档中第一个表格的内部边框。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActiveDocument.Tables.Item(1).Borders.InsideLineStyle = wdLineStyleNone Application.ActiveDocument.Tables.Item(1).Borders.InsideLineStyle = false` |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动文档的第一个表格中的行列之间添加边框。*/ function test() { 	if(Application.ActiveDocument.Tables.Count >= 1) { 	    let tableTemp = ActiveDocument.Tables.Item(1) 	    tableTemp.Borders.InsideLineStyle = true 	} }  /*本示例在活动文档的前 4 个段落之间添加边框。*/ function test() { 	let docActive = Application.ActiveDocument 	let rngTemp = docActive .Range(docActive.Paragraphs.Item(1).Range.Start, docActive.Paragraphs.Item(4).Range.End) 	let Bor = rngTemp.Borders 	Bor.InsideLineStyle = wdLineStyleSingle 	Bor.InsideLineWidth = wdLineWidth150pt } ` |

#### **Borders.InsideLineWidth**

返回或设置对象内部边框的线条宽度。

**语法**

**express.InsideLineWidth**

*express*   一个代表 **Borders** 对象的变量。

**说明**

如果对象的内部边框具有多个线条宽度，则该属性返回 **wdUndefined**；否则，返回 **False** 或 **WdLineWidth** 常量。可将该属性值设置为 **True**、**False** 或下列 **WdLineWidth** 常量之一。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动文档的第一个表格中的行列之间添加边框。*/ function test() { 	if(Application.ActiveDocument.Tables.Count >= 1) { 	    let tableTemp = Application.ActiveDocument.Tables.Item(1) 	    tableTemp.Borders.InsideLineStyle = wdLineStyleDot 	    tableTemp.Borders.InsideLineWidth = wdLineWidth050pt 	} }  /*本示例在活动文档中的前 4 段之间添加点线型边框。*/ function test() { 	let docActive = Application.ActiveDocument 	let rngTemp = docActive.Range(docActive.Paragraphs.Item(1).Range.Start, docActive.Paragraphs.Item(4).Range.End) 	rngTemp.Borders.InsideLineStyle = wdLineStyleDot 	rngTemp.Borders.InsideLineWidth = wdLineWidth075pt }` |

#### **Borders.JoinBorders**

如果删除段落和表格边缘的垂直边框，以使水平边框与页面边框相连，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.JoinBorders**

*express*   一个代表 **Borders** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为所选内容的第一节中的每页添加边框。本示例还删除段落和表格边缘的垂直边框，以使水平边框能与页面边框相连。*/ function test() { 	let Sec = Application.Selection.Sections.Item(1) 	for(let i = 1; i <= Sec.Borders.Count; i++) { 	    Sec.Borders.Item(i).ArtStyle = wdArtBalloonsHotAir 	    Sec.Borders.Item(i).ArtWidth = 15 	} 	let Bor = Sec.Borders 	Bor.DistanceFromLeft = 2 	Bor.DistanceFromRight = 2 	Bor.DistanceFrom = wdBorderDistanceFromText 	Bor.JoinBorders = true }` |

#### **Borders.OutsideColor**

返回或设置外部边框的 24 位颜色。

**语法**

**express.OutsideColor**

*express*   一个代表 **Borders** 对象的变量。

**说明**

此属性可以是任何有效的 **WdColor** 常量，也可以是 Visual Basic 的 **RGB** 函数返回的值。

如果将 **OutsideLineStyle** 属性设置为 **wdLineStyleNone** 或 **False**，则本属性的设置不会产生任何效果。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动文档的第一个表格的行列之间添加边框，然后设置内部和外部边框的颜色。*/ function test() { 	if(Application.ActiveDocument.Tables.Count >= 1) { 	    let myTable = ActiveDocument.Tables.Item(1) 	    let Bor = myTable.Borders 	    Bor.InsideLineStyle = true 	    Bor.InsideColor = wdColorBrightGreen 	    Bor.OutsideColor = wdColorDarkTeal 	} }  /*本示例在活动文档的第一段的周围添加深红色、0.75 磅宽的双边框。*/ function test() { 	let Bor = Application.ActiveDocument.Paragraphs.Item(1).Borders 	Bor.OutsideLineStyle = wdLineStyleDouble 	Bor.OutsideLineWidth = wdLineWidth075pt 	Bor.OutsideColor = wdColorDarkRed }` |

#### **Borders.OutsideColorIndex**

返回或设置外部边框的颜色。**WdColorIndex** 类型，可读写。

**语法**

**express.OutsideColorIndex**

*express*   一个代表 **Borders** 对象的变量。

**说明**

如果将 **OutsideLineStyle** 属性设置为 **wdLineStyleNone** 或 **False**，则该属性的设置不会产生任何效果。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动文档的第一个表格的行列之间添加边框，然后设置内部和外部边框的颜色。*/ function test() { 	if(Application.ActiveDocument.Tables.Count >= 1) { 	    let myTable = Application.ActiveDocument.Tables.Item(1) 	    let Bor = myTable.Borders 	    Bor.InsideLineStyle = true 	    Bor.InsideColorIndex = wdBrightGreen 	    Bor.OutsideColorIndex = wdPink 	} }  /*本示例为活动文档的第一段添加红色、0.75 磅的双边框。*/ function test() { 	let Bor = Application.ActiveDocument.Paragraphs.Item(1).Borders 	Bor.OutsideLineStyle = wdLineStyleDouble 	Bor.OutsideLineWidth = wdLineWidth075pt 	Bor.OutsideColorIndex = wdRed } ` |

#### **Borders.OutsideLineStyle**

返回或设置指定对象的外部边框。

**语法**

**express.OutsideLineStyle**

*express*   一个代表 **Borders** 对象的变量。

**说明**

如果指定对象应用了多种边框，则该属性返回 **wdUndefined**；否则，返回 **False** 或 **WdLineStyle** 常量。可将该属性值设置为 **True**、**False** 或 **WdLineStyle** 常量。

如果为 **True**，则会将线条样式设置为默认线条样式，将线条宽度设置为默认线条宽度。可用 **DefaultBorderLineWidth** 和 **DefaultBorderLineStyle** 属性设置默认线条样式和宽度。

使用下列两条指令中的任何一条，都可删除活动文档中第一个表格的外部边框。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActiveDocument.Tables.Item(1).Borders.OutsideLineStyle = wdLineStyleNone Application.ActiveDocument.Tables.Item(1).Borders.OutsideLineStyle = false` |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动文档的第一段的周围添加 0.75 磅宽的双边框。*/ function test() { 	let Bor = Application.ActiveDocument.Paragraphs.Item(1).Borders 	Bor.OutsideLineStyle = wdLineStyleDouble 	Bor.OutsideLineWidth = wdLineWidth075pt }  /*本示例在活动文档的第一个表格周围添加边框。*/ function test() { 	if(Application.ActiveDocument.Tables.Count >= 1) { 	    let myTable = Application.ActiveDocument.Tables.Item(1) 	    myTable.Borders.OutsideLineStyle = wdLineStyleSingle 	} } ` |

#### **Borders.OutsideLineWidth**

返回或设置对象外部边框的线条宽度。可读写。

**语法**

**express.OutsideLineWidth**

*express*   一个代表 **Borders** 对象的变量。

**说明**

如果对象的外部边框具有多个线条宽度，则该属性返回 **wdUndefined**；否则，返回 **False** 或 **WdLineWidth** 常量。可将该属性值设置为 **True**、**False** 或 **WdLineWidth** 常量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为活动文档的第一个表格添加波浪形边框。*/ function test() { 	if(Application.ActiveDocument.Tables.Count >= 1) { 	    let Bor = Application.ActiveDocument.Tables.Item(1).Borders 	    Bor.OutsideLineStyle = wdLineStyleSingleWavy 	    Bor.OutsideLineWidth = wdLineWidth075pt 	} }  /*本示例在活动文档的前 4 段四周添加点线型边框。*/ function test() { 	let myDoc = Application.ActiveDocument 	let myRange = myDoc.Range(myDoc.Paragraphs.Item(1).Range.Start, myDoc.Paragraphs.Item(4).Range.End) 	myRange.Borders.OutsideLineStyle = wdLineStyleDot 	myRange.Borders.OutsideLineWidth = wdLineWidth075pt }` |

#### **Borders.Parent**

返回一个 **Object** 类型的值，该值代表指定 **Borders** 集合中的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Borders** 对象的变量。

#### **Borders.Shadow**

如果为 **True**，则指定的边框设置为阴影格式。**Boolean** 类型，可读写。

**语法**

**express.Shadow**

*express*   一个代表 **Borders** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在新文档中演示两种不同的边框样式。*/ function test() { 	let myRange = Application.Documents.Add().Content 	myRange.InsertAfter("Demonstration of border with shadow.") 	myRange.InsertParagraphAfter() 	myRange.InsertParagraphAfter() 	myRange.InsertAfter("Demonstration of border without shadow.") 	ActiveDocument.Paragraphs.Item(1).Borders.Shadow = true 	ActiveDocument.Paragraphs.Item(3).Borders.Enable = true } ` |

#### **Borders.SurroundFooter**

如果页面边框环绕文档的页脚，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.SurroundFooter**

*express*   一个代表 **Borders** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置活动文档中第一节的页面边框格式，使页面边框环绕该节中每个页面的页眉和页脚。*/ function test() { 	let Bor = Application.ActiveDocument.Sections.Item(1).Borders 	Bor.SurroundFooter = true 	Bor.SurroundHeader = true }  /*本示例为第一节中的每一页添加艺术型页面边框。该页面边框不环绕页眉和页脚。*/ function test() { 	let Sec = Application.ActiveDocument.Sections.Item(1) 	Sec.Borders.SurroundFooter = false 	Sec.Borders.SurroundHeader = false 	for(let i = 1; i <= Sec.Borders.Count; i++) { 	    Sec.Borders.Item(i).ArtStyle = wdArtPeople 	    Sec.Borders.Item(i).ArtWidth = 15 	} }` |

#### **Borders.SurroundHeader**

如果页面边框环绕文档页眉，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.SurroundHeader**

*express*   一个代表 **Borders** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置活动文档中第一节的页面边框格式，使页面边框不环绕页眉和页脚。*/ function test() { 	let Bor = Application.ActiveDocument.Sections.Item(1).Borders 	Bor.SurroundFooter = false 	Bor.SurroundHeader = false } ` |

适用环境：web

适用平台：windows/linux