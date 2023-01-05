#### **Table**



代表单个表格。**Table** 对象是 **Tables** 集合的一个成员。**Tables** 集合包含指定的所选内容、范围或文档中的所有表格。

**说明**

使用 **Tables**(*Index*) 可返回一个 **Table** 对象，其中 *Index* 为索引号。索引号代表表格在所选内容、范围或文档中的位置。以下示例将活动文档中的第一个表格转换为文本。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActiveDocument.Tables.Item(1).ConvertToText(wdSeparateByTabs)` |

使用 

Add

 方法可在指定范围添加表格。以下示例在活动文档的开头添加一个 3x4 的表格。

 

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `let myRange = Application.ActiveDocument.Range(0, 0) Application.ActiveDocument.Tables.Add(myRange, 3, 4)` |

**方法**

|                                                              | 名称                           | 说明                                                         |
| ------------------------------------------------------------ | ------------------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ApplyStyleDirectFormatting** | 应用指定的样式，但保留用户直接应用的所有格式。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AutoFitBehavior**            | 确定 WPS 如何使用自动调整功能来调整表格的大小。              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AutoFormat**                 | 将预定义外观应用于表格。                                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Cell**                       | 返回一个 **Cell** 对象，该对象代表表格中的一个单元格。       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Converttotext**              | 将表格转换为文本并返回一个 **Range** 对象，该对象代表带分隔符的文本。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**                     | 删除指定的表格。                                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Select**                     | 选择指定的表格。                                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Sort**                       | 对指定的表格进行排序。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SortAscending**              | 按字母数字升序对段落或表格行进行排序。                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SortDescending**             | 以字母数字降序排列表格行。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Split**                      | 在表格中紧靠指定行的上面插入一空段落，并且返回一个 **Table** 对象，此对象包含指定行及其下一行。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **UpdateAutoFormat**           | 使用预定义表格式对表格式进行更新。                           |

**属性**

|                                                              | 名称                      | 说明                                                         |
| ------------------------------------------------------------ | ------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Allowautofit**          | 使 WPS 可以自动调整表格中的单元格的大小以适应内容。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AllowPageBreaks**       | 允许 WPS 对指定表格进行跨页断行。**Boolean** 类型，可读写。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**           | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ApplyStyleColumnBands** | 返回或设置 **Boolean** 值，该值表示如果所应用的预置表样式为列提供了样式带，则是否对表中的列应用样式带。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ApplyStyleFirstColumn** | 如果为 **True**，则 WPS 对指定表格的第一列应用第一列的格式。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ApplyStyleHeadingRows** | 如果为 **True**，则 WPS 为选定表格的第一行应用标题行的格式。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ApplyStyleLastColumn**  | 如果为 **True**，则 WPS 对指定表格的最后一列应用最后一列的格式。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ApplyStyleLastRow**     | 如果为 **True**，则 WPS 对指定表格的最后一行应用最后一行的格式。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ApplyStyleRowBands**    | 返回或设置 **Boolean** 值，该值表示如果所应用的预置表样式为行提供了样式带，则是否对表中的行应用样式带。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatType**        | 返回指定表格所应用的自动套用格式类型。**Long** 类型，只读。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Borders**               | 返回一个 **Borders** 集合，该集合代表指定对象的所有边框。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BottomPadding**         | 返回或设置在表格中单独的单元格或所有单元格内容的下方添加的间距（以磅为单位）。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Columns**               | 返回一个 **Columns** 集合，该集合代表表格中的所有列。只读。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**               | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Descr**                 | 返回或设置包含指定表格的说明的 **String** 类型值。可读/写。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ID**                    | 在将文档另存为网页时，返回或设置指定的表的标识标签。可读/写 **String** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LeftPadding**           | 返回或设置在表格中所有单元格的内容左侧要增加的间距（以磅为单位）。可读写 **Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Nestinglevel**          | 返回指定表格的嵌套层。只读 **Long** 类型。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                | 返回一个 **Object** 类型值，该值代表指定 **Table** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Preferredwidth**        | 返回或设置指定表格的指定宽度（以磅为单位或以窗口宽度的百分比表示）。可读写 **Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Preferredwidthtype**    | 返回指定表格所应用的自动套用格式类型。**Long** 类型，只读。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Range**                 | 返回一个 **Range** 对象，该对象代表指定表格中所含的文档部分。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RightPadding**          | 返回或设置在表格的所有单元格的内容右侧要增加的间距（以磅为单位）。可读写 **Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Rows**                  | 返回一个 **Rows** 集合，该集合代表表格中所有的表格行。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shading**               | 返回一个 **Shading** 对象，该对象代表指定对象的底纹格式。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Spacing**               | 返回或设置表格中单元格的间距（以磅为单位）。可读写 **Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Style**                 | 返回或设置指定表格的样式。可读写 **Variant** 类型。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TableDirection**        | 返回或设置 WPS 对指定表格中的单元格进行排序的方向。可读写 **WdTableDirection** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Tables**                | 返回一个 **Tables** 集合，该集合代表指定表格中的所有嵌套表格。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Title**                 | 返回或设置包含指定表格的标题的 **String** 类型值。可读/写。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TopPadding**            | 返回或设置表格中所有单元格的内容上方要增加的间距（以磅为单位）。可读写 **Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Uniform**               | 如果表格中所有的行具有相同的列数，则该属性值为 **True**。**Boolean** 类型，只读。 |

**成员方法**

#### **Table.ApplyStyleDirectFormatting**

应用指定的样式，但保留用户直接应用的所有格式。

**语法**

**express.ApplyStyleDirectFormatting(StyleName)**

*express*   一个代表 **Table** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**             |
| ----------- | ------------- | ------------ | -------------------- |
| *StyleName* | 必选          | **String**   | 要应用的样式的名称。 |

#### **Table.AutoFitBehavior**

确定 WPS 如何使用自动调整功能来调整表格的大小。

**语法**

**express.AutoFitBehavior(Behavior)**

*express*   一个代表 **Table** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型**          | **说明**                                                |
| ---------- | ------------- | --------------------- | ------------------------------------------------------- |
| *Behavior* | 必选          | **WdAutoFitBehavior** | 指定 WPS 如何使用“自动调整”功能重新调整指定表格的大小。 |

**说明**

WPS 可以根据表格单元格的内容或文档窗口的宽度重新调整表格的大小。也可使用本方法来关闭“自动调整”功能。这样，表格的大小是固定值，而不随单元格内容或窗口宽度而改变。

将 **AutoFitBehavior** 属性设置为 **wdAutoFitContent** 或 **wdAutoFitWindow** 会将 **AllowAutoFit** 属性设置为 **True**（如果该属性当前为 **False**）。同样，将 **AutoFitBehavior** 属性设置为 **wdAutoFitFixed** 会将 **AllowAutoFit** 属性设置为 **False**（如果该属性当前为 **True**）。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：设定活动文档中第一张表格的“自动调整”操作，使其可根据文档窗口的宽度自动调整大小。*/ Application.ActiveDocument.Tables.Item(1).AutoFitBehavior(wdAutoFitWindow)` |

#### **Table.AutoFormat**

将预定义外观应用于表格。

**语法**

**express.AutoFormat(Format, ApplyBorders, ApplyShading, ApplyFont, ApplyColor, ApplyHeadingRows, ApplyLastRow, ApplyFirstColumn, ApplyLastColumn, AutoFit)**

*express*   一个代表 **Table** 对象的变量。

**参数**

| **名称**           | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------------ | ------------- | ------------ | ------------------------------------------------------------ |
| *Format*           | 可选          | **Variant**  | 要应用的格式。此参数可以是 WdTableFormat 常量、WdTableFormatApply 常量或 TableStyle 对象。 |
| *ApplyBorders*     | 可选          | **Variant**  | 如果该参数值为 True，则应用指定格式的边框属性。默认值为 True。 |
| *ApplyShading*     | 可选          | **Variant**  | 如果该参数值为 True，则应用指定格式的底纹属性。默认值为 True。 |
| *ApplyFont*        | 可选          | **Variant**  | 如果该参数值为 True，则应用指定格式的字体属性。默认值为 True。 |
| *ApplyColor*       | 可选          | **Variant**  | 如果该参数值为 True，则应用指定格式的颜色属性。默认值为 True。 |
| *ApplyHeadingRows* | 可选          | **Variant**  | Variant 如果该参数值为 True，则应用指定格式的标题行属性。默认值为 True。 |
| *ApplyLastRow*     | 可选          | **Variant**  | 如果该参数值为 True，则应用指定格式的最后一行的属性。默认值为 False。 |
| *ApplyFirstColumn* | 可选          | **Variant**  | 如果该参数值为 True，则应用指定格式的第一列的属性。默认值为 True。 |
| *ApplyLastColumn*  | 可选          | **Variant**  | 如果该参数值为 True，则应用指定格式的最后一列的属性。默认值为 False。 |
| *AutoFit*          | 可选          | **Variant**  | 如果该参数值为 True，则在不改变单元格内文字的换行方式的前提下尽可能缩小表格列宽。默认值为 True。 |

**说明**

此方法的参数对应于**“表格自动套用格式”**对话框中的选项。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在新文档中创建一个 5x5 的表格，并将“彩色型 2”格式的所有属性应用于此表格。*/ function test() { let newDoc = Documents.Add() let myTable = newDoc.Tables.Add(Selection.Range, 5, 5) myTable.AutoFormat(wdTableFormatColorful2) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将“古典型 2”格式的所有属性应用于插入点所在表格。如果插入点不在表格中，则显示一个消息框。*/ function test() { Selection.Collapse(wdCollapseStart)  if(Selection.Information(wdWithInTable) == true){     Selection.Tables.Item(1).AutoFormat(wdTableFormatClassic2) } else{     MsgBox("The insertion point is not in a table.") } }` |

#### **Table.Cell**

返回一个 **Cell** 对象，该对象代表表格中的一个单元格。

**语法**

**express.Cell(Row, Column)**

*express*   一个代表 **Table** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                  |
| -------- | ------------- | ------------ | --------------------------------------------------------- |
| *Row*    | 必选          | **Long**     | 要返回的表格行数。可以是介于 1 和表格行数之间的任意整数。 |
| *Column* | 必选          | **Long**     | 要返回的表格列数。可以是介于 1 和表格列数之间的任意整数。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在新文档中创建一个 3x3 表格，并在表格的第一个和最后一个单元格中插入文本*/ function test() {     let docNew = Application.Documents.Add()     let tableNew = docNew.Tables.Add(Selection.Range, 3, 3)     	     tableNew.Cell(1,1).Range.InsertAfter("First cell")     tableNew.Cell(tableNew.Rows.Count, tableNew.Columns.Count).Range.InsertAfter("Last Cell") }  /*以下示例删除活动文档的第一个表格中的第一个单元格*/ function test() {     if(Application.ActiveDocument.Tables.Count >= 1){         Application.ActiveDocument.Tables.Item(1).Cell(1, 1).Delete()     } }` |

#### **Table.Converttotext**

将表格转换为文本并返回一个 **Range** 对象，该对象代表带分隔符的文本。

**语法**

**express.Converttotext(Separator, NestedTables )**

*express*   一个代表 **Table** 对象的变量。

**参数**

| **名称**       | **必选/可选** | **数据类型** | **说明**                                                     |
| -------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Separator*    | 可选          | **Variant**  | 用以分隔被转换列的分隔符（被转换行由段落标记分隔）。可以是下列 **WdTableFieldSeparator** 常量之一。 |
| *NestedTables* | 可选          | **Variant**  | 如果要将嵌套的表格转换为文本，则为 **True**。如果 **Separator** 不是 **wdSeparateByParagraphs**，则将忽略此参数。默认值为 True。 |

**说明**

将 **ConvertToText** 方法应用于一个 **Table** 对象时，该对象将被删除。如果要保留对已转换的表格内容的引用，就必须为 **ConvertToText** 方法返回的 **Range** 对象赋予新的对象变量。在下面示例中，将活动文档第一个表格转换为文本，并将其格式设为项目符号列表。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { 	let tableTemp = Application.ActiveDocument.Tables.Item(1) 	let rngTemp = tableTemp.ConvertToText(wdSeparateByParagraphs)  	rngTemp.ListFormat.ApplyListTemplate(ListGalleries.Item(wdBulletGallery).ListTemplates.Item(1)) }` |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例创建一张表格，然后将其转换为文本，以制表符作为分隔字符*/ function test() { 	let docNew = Application.Documents.Add() 	let tableNew = docNew.Tables.Add(Selection.Range, 3, 3) 	 	for(let i = 1; i <= tableNew.Range.Cells.Count; i++){ 	    tableNew.Range.Cells.Item(i).Range.InsertAfter("Cell" + i) 	} 	MsgBox("Click OK to convert table to text.") 	let rngTemp = tableNew.ConvertToText(wdSeparateByTabs) }  /*本示例将包含选定内容的表格转换为文本，各列之间用空格分隔*/ function test() { 	if(Application.Selection.Information(wdWithInTable) == true){     	    Application.Selection.Tables.Item(1).ConvertToText(" ") 	} 	else{ 	    MsgBox("The insertion point is not in a table.") 	} }` |

#### **Table.Delete**

删除指定的表格。

**语法**

**express.Delete()**

*express*   一个代表 **Table** 对象的变量。

#### **Table.Select**

选择指定的表格。

**语法**

**express.Select()**

*express*   一个代表 **Table** 对象的变量。

**说明**

使用此方法后，可使用 **Selection** 对象处理所选项目。有关详细信息，请参阅 处理 Selection 对象。

#### **Table.Sort**

对指定的表格进行排序。

**语法**

**express.Sort(ExcludeHeader, FieldNumber, SortFieldType, SortOrder, FieldNumber2, SortFieldType2, SortOrder2, FieldNumber3, SortFieldType3, SortOrder3, CaseSensitive, BidiSort, IgnoreThe, IgnoreKashida, IgnoreDiacritics , IgnoreHe, LanguageID)**

*express*   一个代表 **Table** 对象的变量。

**参数**

| **名称**           | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------------ | ------------- | ------------ | ------------------------------------------------------------ |
| *ExcludeHeader*    | 可选          | **Variant**  | 如果该参数值为 **True**，则不对首行进行排序。默认值为 **False**。 |
| *FieldNumber*      | 可选          | **Variant**  | 用于排序的第一个域。WPS 依次按 **FieldNumber**、**FieldNumber2、FieldNumber3 排序。** |
| *SortFieldType*    | 可选          | **Variant**  | **FieldNumber** 的排序类型。可以是 **WdSortFieldType** 常量之一。根据您选择或安装的语言支持（例如，美国英语），以上某些常量可能不可用。默认值为 **wdSortFieldAlphanumeric**。 |
| *SortOrder*        | 可选          | **Variant**  | 对 **FieldNumber 进行排序时使用的排序顺序。可以是 WdSortOrder 常量。** |
| *FieldNumber2*     | 可选          | **Variant**  | 用于排序的第二个域。                                         |
| *SortFieldType2*   | 可选          | **Variant**  | **FieldNumber2** 的排序类型。可以是 **WdSortFieldType** 常量之一。根据您选择或安装的语言支持（例如，美国英语），以上某些常量可能不可用。默认值为 **wdSortFieldAlphanumeric**。 |
| *SortOrder2*       | 可选          | **Variant**  | 对 **FieldNumber2** 进行排序时使用的排序顺序。可以是一个 **WdSortOrder** 常量。 |
| *FieldNumber3*     | 可选          | **Variant**  | 用于排序的第三个域。                                         |
| *SortFieldType3*   | 可选          | **Variant**  | **FieldNumber3** 的排序类型。可以是 **WdSortFieldType** 常量之一。根据您选择或安装的语言支持（例如，美国英语），以上某些常量可能不可用。默认值为 **wdSortFieldAlphanumeric**。 |
| *SortOrder3*       | 可选          | **Variant**  | 对 **FieldNumber3** 进行排序时使用的排序顺序。可以是一个 **WdSortOrder** 常量。 |
| *CaseSensitive*    | 可选          | **Variant**  | 如果该参数值为 **True**，则排序时区分大小写。默认值为 **False**。 |
| *BidiSort*         | 可选          | **Variant**  | 如果该参数值为 **True，则基于从右向左语言的规则进行排序。根据您所选择或安装的语言支持（例如，美国英语），此参数可能不可用。** |
| *IgnoreThe*        | 可选          | **Variant**  | 如果该参数值为 **True，则对从右向左的语言文本进行排序时忽略阿拉伯字符 alef lam。根据您所选择或安装的语言支持（例如，美国英语），此参数可能不可用。** |
| *IgnoreKashida*    | 可选          | **Variant**  | 如果该参数值为 **True**，则对从右向左的语言文本进行排序时忽略 **Kashidas**。根据您所选择或安装的语言支持（例如，美国英语），此参数可能不可用。 |
| *IgnoreDiacritics* | 可选          | **Variant**  | 如果该参数值为 **True**，则对从右向左的语言文本进行排序时忽略双向控制字符。根据您所选择或安装的语言支持（例如，美国英语），此参数可能不可用。 |
| *IgnoreHe*         | 可选          | **Variant**  | 如果该参数值为 **True**，则对从右向左的语言文本进行排序时忽略希伯来字符 **he**。根据您所选择或安装的语言支持（例如，美国英语），此参数可能不可用。 |
| *LanguageID*       | 可选          | **Variant**  | 指定排序语言。可以是 **WdLanguageID** 常量之一。要获得 **WdLanguageID** 常量的列表，请参阅“对象浏览器”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例对活动文档中的第一个表格进行排序，首行除外*/ function NewTableSort(){     Application.ActiveDocument.Tables.Item(1).Sort(true) }` |

#### **Table.SortAscending**

按字母数字升序对段落或表格行进行排序。

**语法**

**express.SortAscending()**

*express*   一个代表 **Table** 对象的变量。

**说明**

将表格第一行视为域名记录，不进行排序。可用 **Sort** 方法将首行包括在排序中。此方法提供针对包含数据列的邮件合并数据源进行排序的一种简化形式。 对于大多数排序任务，请使用 **Sort** 方法。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例以升序排列包含所选内容的表格。*/ function test() { if(Selection.Information(wdWithInTable) == true){     Selection.Tables.Item(1).SortAscending() } else{     MsgBox("The insertion point is not in a table.") } }` |

#### **Table.SortDescending**

以字母数字降序排列表格行。

**语法**

**express.SortDescending()**

*express*   一个代表 **Table** 对象的变量。

**说明**

将表格第一行视为域名记录，不进行排序。使用 **Sort** 方法将域名记录包含在排序中。

此方法提供针对包含数据列的邮件合并数据源进行排序的一种简化形式。对于大多数排序任务，请使用 **Sort** 方法。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在新文档中创建一张 5x5 表格，将文本插入每个单元格，然后以字母数字降序排列表格。*/ function test() { let newDoc = Documents.Add() let myTable = newDoc.Tables.Add(Selection.Range, 5, 5)  for(let i = 1; i <= myTable.Rows.Count; i++){     for(let j = 1; j <= myTable.Columns.Count; j++){         let MyRange = myTable.Rows.Item(i).Cells.Item(j).Range         MyRange.InsertAfter("Cell " + i + "," + j)     } } MsgBox("Click OK to sort in descending order.") myTable.SortDescending() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例以字母数字降序排列插入点所在的表格。*/ function test() { if(Selection.Information(wdWithInTable) == true){     Selection.Tables.Item(1).SortDescending() } else{     MsgBox("The insertion point is not in a table.") } }  ` |

#### **Table.Split**

在表格中紧靠指定行的上面插入一空段落，并且返回一个 **Table** 对象，此对象包含指定行及其下一行。

**语法**

**express.Split(BeforeRow)**

*express*   一个代表 **Table** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                        |
| ----------- | ------------- | ------------ | ----------------------------------------------- |
| *BeforeRow* | 必选          | **Variant**  | 将要拆分的表格的前一行。可以是行号或 Row 对象。 |

#### **Table.UpdateAutoFormat**

使用预定义表格式对表格式进行更新。

**语法**

**express.UpdateAutoFormat()**

*express*   一个代表 **Table** 对象的变量。

**说明**

下面举例说明本方法的作用：如果使用 **AutoFormat** 来应用一种表格式，然后插入行和列，那么表格的外观可能不再像预定义的那样。**UpdateAutoFormat** 会恢复该格式。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例创建一个表格并应用预定义格式，再添加一行，然后重新应用预定义格式。*/ function test() { let docNew = Documents.Add() let tableNew = docNew.Tables.Add(Selection.Range, 5, 5)  tableNew.AutoFormat(wdTableFormatColumns1) tableNew.Rows.Add(tableNew.Rows.Item(1))  MsgBox("Click OK to reapply autoformatting.") tableNew.UpdateAutoFormat() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例恢复插入点所在表格的预定义格式。*/ function test() { if(Selection.Information(wdWithInTable) == true){     Selection.Tables.Item(1).UpdateAutoFormat() } else{     MsgBox("The insertion point is not in a table.") } }` |

**成员属性**

#### **Table.Allowautofit**

使 WPS 可以自动调整表格中的单元格的大小以适应内容。**Boolean** 类型，可读写。

**语法**

**express.Allowautofit**

*express*   一个代表 **Table** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置活动文档的第一张表格可自动调整大小以适应内容*/ function AllowFit(){     Applicaiton.ActiveDocument.Tables.Item(1).AllowAutoFit = true }` |

#### **Table.AllowPageBreaks**

允许 WPS 对指定表格进行跨页断行。**Boolean** 类型，可读写。

**语法**

**express.AllowPageBreaks**

*express*   一个代表 **Table** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置活动文档的第二个表格跨页断行。*/ function test(){     ActiveDocument.Tables.Item(2).AllowPageBreaks = true }` |

#### **Table.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **Table** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **Table.ApplyStyleColumnBands**

返回或设置 **Boolean** 值，该值表示如果所应用的预置表样式为列提供了样式带，则是否对表中的列应用样式带。可读/写。

**语法**

**express.ApplyStyleColumnBands**

*express*   一个代表 **Table** 对象的变量。

#### **Table.ApplyStyleFirstColumn**

如果为 **True**，则 WPS 对指定表格的第一列应用第一列的格式。**Boolean** 类型，可读写。

**语法**

**express.ApplyStyleFirstColumn**

*express*   一个代表 **Table** 对象的变量。

**说明**

指定的表格样式必须包含第一列的格式，才能将此格式应用于表格。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例用“Table Style 1”的表格样式设置活动文档中的第二个表格的格式，并删除第一行和最后一行、第一列和最后一列的格式。本示例假定名为“Table Style 1”的表格样式存在，并且包含第一列的格式。*/ function test(){     let table = ActiveDocument.Tables.Item(2)      table.Style = "Table Style 1"     table.ApplyStyleFirstColumn = false     table.ApplyStyleHeadingRows = false     table.ApplyStyleLastColumn = false     table.ApplyStyleLastRow = false }` |

#### **Table.ApplyStyleHeadingRows**

如果为 **True**，则 WPS 为选定表格的第一行应用标题行的格式。**Boolean** 类型，可读写。

**语法**

**express.ApplyStyleHeadingRows**

*express*   一个代表 **Table** 对象的变量。

**说明**

指定的表格样式必须包含标题行的格式，才能将此格式应用于表格。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例用“Table Style 1”的表格样式设置活动文档中的第二个表格的格式，并删除第一行和最后一行、第一列和最后一列的格式。本示例假定名为“Table Style 1”的表格样式存在，并且包含标题行的格式。*/ function TableStyles(){     let table = ActiveDocument.Tables.Item(2)      table.Style = "Table Style 1"     table.ApplyStyleFirstColumn = false     table.ApplyStyleHeadingRows = false     table.ApplyStyleLastColumn = false     table.ApplyStyleLastRow = false }` |

#### **Table.ApplyStyleLastColumn**

如果为 **True**，则 WPS 对指定表格的最后一列应用最后一列的格式。**Boolean** 类型，可读写。

**语法**

**express.ApplyStyleLastColumn**

*express*   一个代表 **Table** 对象的变量。

**说明**

指定的表格样式必须包含最后一列的格式，才能将此格式应用于表格。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例用“Table Style 1”的表格样式设置活动文档中的第二个表格的格式，并删除第一行和最后一行、第一列和最后一列的格式。本示例假定名为“Table Style 1”的表格样式存在，并且其包含最后一列的格式。*/ function TableStyles(){     let table = ActiveDocument.Tables.Item(2)      table.Style = "Table Style 1"     table.ApplyStyleFirstColumn = false     table.ApplyStyleHeadingRows = false     table.ApplyStyleLastColumn = false     table.ApplyStyleLastRow = false }` |

#### **Table.ApplyStyleLastRow**

如果为 **True**，则 WPS 对指定表格的最后一行应用最后一行的格式。**Boolean** 类型，可读写。

**语法**

**express.ApplyStyleLastRow**

*express*   一个代表 **Table** 对象的变量。

**说明**

指定的表格样式必须包含最后一行的格式，才能将此格式应用于表格。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例用“Table Style 1”的表格样式设置活动文档中的第二个表格的格式，并删除第一行和最后一行、第一列和最后一列的格式。本示例假定名为“Table Style 1”的表格样式存在，并且其包含最后一行的格式。*/ function test(){     let table = ActiveDocument.Tables.Item(2)      table.Style = "Table Style 1"     table.ApplyStyleFirstColumn = false     table.ApplyStyleHeadingRows = false     table.ApplyStyleLastColumn = false     table.ApplyStyleLastRow = false }` |

#### **Table.ApplyStyleRowBands**

返回或设置 **Boolean** 值，该值表示如果所应用的预置表样式为行提供了样式带，则是否对表中的行应用样式带。可读/写。

**语法**

**express.ApplyStyleRowBands**

*express*   一个代表 **Table** 对象的变量。

#### **Table.AutoFormatType**

返回指定表格所应用的自动套用格式类型。**Long** 类型，只读。

**语法**

**express.AutoFormatType**

*express*   一个代表 **Table** 对象的变量。

**说明**

该属性可以是 **WdTableFormat** 常量之一。可使用 **AutoFormat** 方法为表格设置自动套用格式。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果活动文档的第一个表格的当前格式是“简明型 1”、“简明型 2”或“简明型 3”，那么本示例为该表格自动套用“古典型 1”格式。*/ function test() { if(ActiveDocument.Tables.Count >= 1){     if(ActiveDocument.Tables.Item(1).AutoFormatType <= wdTableFormatSimple3){         ActiveDocument.Tables.Item(1).AutoFormat(wdTableFormatClassic1)     } } }` |

#### **Table.Borders**

返回一个 **Borders** 集合，该集合代表指定对象的所有边框。

**语法**

**express.Borders**

*express*   一个代表 **Table** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对活动文档中的第一个表格应用内部和外部边框*/ function test(){     let myTable = Application.ActiveDocument.Tables.Item(1)     myTable.Borders.InsideLineStyle = wdLineStyleSingle     myTable.Borders.OutsideLineStyle = wdLineStyleDouble }` |

#### **Table.BottomPadding**

返回或设置在表格中单独的单元格或所有单元格内容的下方添加的间距（以磅为单位）。**Single** 类型，可读写。

**语法**

**express.BottomPadding**

*express*   一个代表 **Table** 对象的变量。

**说明**

单独单元格的 **BottomPadding** 属性的设置值会覆盖整个表格的 **BottomPadding** 属性的设置值。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中第一个表格的底边距设为 40 像素。*/ ActiveDocument.Tables.Item(1).BottomPadding = PixelsToPoints(40, true)` |

#### **Table.Columns**

返回一个 **Columns** 集合，该集合代表表格中的所有列。只读。

**语法**

**express.Columns**

*express*   一个代表 **Table** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示活动文档的第一个表格中的列数*/ if(Application.ActiveDocument.Tables.Count >= 1){     alart(Application.ActiveDocument.Tables.Item(1).Columns.Count) }` |

#### **Table.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Table** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **Table.Descr**

返回或设置包含指定表格的说明的 **String** 类型值。可读/写。

**语法**

**express.Descr**

*express*   一个代表 **Table** 对象的变量。

**说明**

使用 **Descr** 属性可提供表格的可选文字说明。该属性将文本添加到 WPS 2015 中**“表格属性”**对话框的**“可选文字”**选项卡上的**“说明”**文本框中。

| 注释                                                         |
| ------------------------------------------------------------ |
| Web 浏览器在加载表格的过程中或表格丢失时显示可选文字。Web 搜索引擎利用可选文字帮助查找网页。可选文字也可用来帮助残障人士。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例向活动文档的第一个表格中添加可选文字表格说明。*/ function test() { let doc = ActiveDocument let tbl = doc.Tables.Item(1) tbl.Descr = "This is a table description." }` |

#### **Table.ID**

在将文档另存为网页时，返回或设置指定的表的标识标签。可读/写 **String** 类型。

**语法**

**express.ID**

*express*   一个代表 **Table** 对象的变量。

#### **Table.LeftPadding**

返回或设置在表格中所有单元格的内容左侧要增加的间距（以磅为单位）。可读写 **Single** 类型。

**语法**

**express.LeftPadding**

*express*   一个代表 **Table** 对象的变量。

**说明**

单个单元格的 **LeftPadding** 属性设置将覆盖整个表格的 **LeftPadding** 属性设置。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中第一个表格的左边距设为 40 像素。*/ ActiveDocument.Tables.Item(1).LeftPadding = PixelsToPoints(40, false)` |

#### **Table.Nestinglevel**

返回指定表格的嵌套层。只读 **Long** 类型。

**语法**

**express.Nestinglevel**

*express*   一个代表 **Table** 对象的变量。

**说明**

最外围表格的嵌套层为 1。每一个相连嵌套表格的嵌套层比其前面表格的嵌套层高 1。

#### **Table.Parent**

返回一个 **Object** 类型值，该值代表指定 **Table** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Table** 对象的变量。

#### **Table.Preferredwidth**

返回或设置指定表格的指定宽度（以磅为单位或以窗口宽度的百分比表示）。可读写 **Single** 类型。

**语法**

**express.Preferredwidth**

*express*   一个代表 **Table** 对象的变量。

**说明**

如果 **PreferredWidthType**属性设置为 **wdPreferredWidthPoints**，则 **PreferredWidth** 属性返回或设置宽度（以磅为单位）。如果 **PreferredWidthType** 属性设置为 **wdPreferredWidthPercent**，则 **PreferredWidth** 属性返回或设置宽度（以窗口宽度的百分比表示）。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将 WPS 设置为以窗口宽度的百分比来表示指定宽度，然后将文档中第一张表格的指定宽度设置为窗口宽度的 50%*/ Application.ActiveDocument.Tables.Item(1).PreferredWidthType = wdPreferredWidthPercent Application.ActiveDocument.Tables.Item(1).PreferredWidth = 50` |

#### **Table.Preferredwidthtype**

返回指定表格所应用的自动套用格式类型。**Long** 类型，只读。

**语法**

**express.Preferredwidthtype**

*express*   一个代表 **Table** 对象的变量。

**说明**

该属性可以是 **WdTableFormat** 常量之一。可使用 **AutoFormat** 方法为表格设置自动套用格式。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果活动文档的第一个表格的当前格式是“简明型 1”、“简明型 2”或“简明型 3”，那么本示例为该表格自动套用“古典型 1”格式*/ if(Application.ActiveDocument.Tables.Count >= 1){     if(Application.ActiveDocument.Tables.Item(1).AutoFormatType <= wdTableFormatSimple3){         Application.ActiveDocument.Tables.Item(1).AutoFormat(wdTableFormatClassic1)     } }` |

#### **Table.Range**

返回一个 **Range** 对象，该对象代表指定表格中所含的文档部分。

**语法**

**express.Range**

*express*   一个代表 **Table** 对象的变量。

#### **Table.RightPadding**

返回或设置在表格的所有单元格的内容右侧要增加的间距（以磅为单位）。可读写 **Single** 类型。

**语法**

**express.RightPadding**

*express*   一个代表 **Table** 对象的变量。

**说明**

单个单元格的 **RightPadding** 属性设置将覆盖整个表格的 **RightPadding** 属性设置。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中第一个表格的右边距设置为 40 像素。*/ ActiveDocument.Tables.Item(1).RightPadding = PixelsToPoints(40, false)` |

#### **Table.Rows**

返回一个 **Rows** 集合，该集合代表表格中所有的表格行。只读。

**语法**

**express.Rows**

*express*   一个代表 **Table** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例删除活动文档第一个表格的第二行*/ Application.ActiveDocument.Tables.Item(1).Rows.Item(2).Delete()` |

#### **Table.Shading**

返回一个 **Shading** 对象，该对象代表指定对象的底纹格式。

**语法**

**express.Shading**

*express*   一个代表 **Table** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例对活动文档中第一个表格应用横线纹理。*/ if(ActiveDocument.Tables.Count >= 1){     ActiveDocument.Tables.Item(1).Shading.Texture = wdTextureHorizontal }` |

#### **Table.Spacing**

返回或设置表格中单元格的间距（以磅为单位）。可读写 **Single** 类型。

**语法**

**express.Spacing**

*express*   一个代表 **Table** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档的第一个表格中的单元格的间距设置为 9 磅。*/ ActiveDocument.Tables.Item(1).Spacing = 9` |

#### **Table.Style**

返回或设置指定表格的样式。可读写 **Variant** 类型。

**语法**

**express.Style**

*express*   一个代表 **Table** 对象的变量。

**说明**

要设置该属性，您可以指定样式的本地名称、一个整数、一个 **WdBuiltinStyle** 常量或一个包含所需表格格式的 **TableStyle** 对象。

#### **Table.TableDirection**

返回或设置 WPS 对指定表格中的单元格进行排序的方向。可读写 **WdTableDirection** 类型。

**语法**

**express.TableDirection**

*express*   一个代表 **Table** 对象的变量。

**说明**

如果将 **TableDirection** 属性设为 **wdTableDirectionLtr**，则选定的行按照最左侧的第一列进行排列。如果将 **TableDirection** 属性设为 **wdTableDirectionRtl**，则选定的行按照最右侧的第一列进行排列。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例设置 WPS 从右向左对选定行中的单元格进行排序。*/ Selection.Rows.TableDirection = wdTableDirectionRtl` |

#### **Table.Tables**

返回一个 **Tables** 集合，该集合代表指定表格中的所有嵌套表格。只读。

**语法**

**express.Tables**

*express*   一个代表 **Table** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

#### **Table.Title**

返回或设置包含指定表格的标题的 **String** 类型值。可读/写。

**语法**

**express.Title**

*express*   一个代表 **Table** 对象的变量。

**说明**

使用 **Title** 属性可为表格提供可选文字标题。该属性将标题文本添加到 WPS 2015 中**“表格属性”**对话框的**“可选文字”**选项卡上的**“标题”**文本框中。

| 注释                                                         |
| ------------------------------------------------------------ |
| Web 浏览器在加载表格的过程中或表格丢失时显示可选文字。Web 搜索引擎利用可选文字帮助查找网页。可选文字也可用来帮助残障人士。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例向活动文档中的第一个表格中添加可选文字表格标题。*/ ActiveDocument.Tables.Item(1).Title = "Table 1."` |

#### **Table.TopPadding**

返回或设置表格中所有单元格的内容上方要增加的间距（以磅为单位）。可读写 **Single** 类型。

**语法**

**express.TopPadding**

*express*   一个代表 **Table** 对象的变量。

**说明**

单个单元格的 **TopPadding** 属性设置将覆盖整个表格的 **TopPadding** 属性设置。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中第一个表格的上边距设为 40 像素。*/ ActiveDocument.Tables.Item(1).TopPadding = PixelsToPoints(40, true)` |

#### **Table.Uniform**

如果表格中所有的行具有相同的列数，则该属性值为 **True**。**Boolean** 类型，只读。

**语法**

**express.Uniform**

*express*   一个代表 **Table** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例创建一个包含拆分单元格的表格，然后显示一个消息框，确认并非表格中的每一行都具有相同的列数。*/ function test() { let newDoc = Documents.Add() let myTable = newDoc.Tables.Add(Selection.Range, 5, 5) myTable.Cell(3, 3).Split(1, 2)  if(!myTable.Uniform){     MsgBox("Table is not uniform") } }  ` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例确定包含选定内容的表格中的各行是否具有相同的列数。*/ function test() { if(Selection.Information(wdWithInTable) == true){     MsgBox(Selection.Tables.Item(1).Uniform) } }` |

适用环境：web

适用平台：windows/linux