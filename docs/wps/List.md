#### **List**



代表已应用于文档中指定段落的单一列表格式。**List** 对象是 **Lists** 集合的一个成员。

**说明**

使用 **Lists**(*Index*) 可返回一个 **List** 对象，其中 *Index* 是索引号。以下示例返回活动文档中第二个列表中的项目数。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActiveDocument.Lists.Item(2).CountNumberedItems()` |

要返回具有列表格式的所有段落，请使用 **ListParagraphs** 属性。要将这些段落作为范围返回，请使用 **Range** 属性。

要对现有列表应用不同的列表格式，请使用 **ApplyListTemplate** 方法和 **List** 对象。要向文档添加新列表，请对指定范围使用 **ApplyListTemplate** 方法和 **ListFormat** 对象。

使用 **CanContinuePreviousList** 方法可确定是否能继续使用以前应用于文档的列表中的列表格式。

使用 **CountNumberedItems** 方法可返回编号列表或项目符号列表中的项目数，包括 LISTNUM 域。

要确定一个列表是否包含多个列表模板，请使用 **SingleListTemplate** 属性。

您可以处理文档中的单个 **List** 对象，但是要更精确地控制，则应使用 **ListFormat** 对象。

注释：图片项目符号列表不包含在 **Lists** 集合中，并且不能使用 **List** 对象进行处理。

**方法**

|                                                              | 名称                        | 说明                                                         |
| ------------------------------------------------------------ | --------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ApplyListTemplate**       | 为指定的 **ListFormat** 对象设置列表格式。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CanContinuePreviousList** | 返回一个 **WdContinue** 常量（**wdContinueDisabled**、**wdResetList** 或 **wdContinueList**），该常量表明是否接续上一列表格式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ConvertNumbersToText**    | 更改指定的 **List** 对象中的列表编号和 LISTNUM 域。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CountNumberedItems**      | 返回指定 **List** 对象中的项目符号或编号项目和 LISTNUM 域的数量。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **RemoveNumbers**           | 删除指定列表中的编号或项目符号。                             |

**属性**

|                                                              | 名称                   | 说明                                                         |
| ------------------------------------------------------------ | ---------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ListParagraphs**     | 返回一个 **ListParagraphs** 集合，该集合代表列表、文档或范围中的所有编号段落。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Range**              | 返回一个 **Range** 对象，该对象代表指定对象所包含的文档部分。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SingleListTemplate** | 如果整个列表都使用同一列表模板，则该属性值为 **True**。只读 **Boolean** 类型。 |

**成员方法**

#### **List.ApplyListTemplate**

为指定的 **ListFormat** 对象设置列表格式。

**语法**

**express.ApplyListTemplate(ListTemplate, ContinuePreviousList, ApplyTo, DefaultListBehavior)**

*express*   一个代表 **List** 对象的变量。

**参数**

| **名称**               | **必选/可选** | **数据类型**     | **说明**                                                     |
| ---------------------- | ------------- | ---------------- | ------------------------------------------------------------ |
| *ListTemplate*         | 必选          | **ListTemplate** | 要应用的列表模板。                                           |
| *ContinuePreviousList* | 可选          | **Variant**      | 如果为 True，则接续前一列表的编号；如果为 False，则新建一个列表。 |
| *ApplyTo*              | 可选          | **Variant**      | 要应用列表模板的列表部分。可以是下列 WdListApplyTo 常量之一：wdListSelection、wdListWholeList 或 wdListThisPointForward。 |
| *DefaultListBehavior*  | 可选          | **Variant**      | 设置可指定 WPS 是否使用新的网页格式，以便列表具有更好的显示效果的值。其值可取下列 WdDefaultListBehavior 常量之一：wdWord8ListBehavior（使用与 WPS 97 兼容的格式）或 wdWord9ListBehavior（使用适于网页的格式）。考虑到兼容问题，默认常量为 wdWord8ListBehavior。但是在新建过程中若要建立缩进式多级列表，应当使用 wdWord9ListBehavior，以利用改进过的适用于网页的格式。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例将变量 myRange 设置为活动文档的一个区域，然后检查该区域，判断其中是否存在列表格式。如果没有列表格式，则为该区域设置第四种符号列表模板的格式 */ function test() {   	let actDoc = Application.ActiveDocument;   	let rng = actDoc.Range(actDoc.Paragraphs.Item(1).Range.Start, actDoc.Paragraphs.Item(2).Range.End);   	if (rng.ListFormat.ListType === wdListNoNumbering) {   		rng.ListFormat.ApplyListTemplate(Application.ListGalleries.Item(2).ListTemplates.Item(4))   	} }` |

#### **List.CanContinuePreviousList**

返回一个 **WdContinue** 常量（**wdContinueDisabled**、**wdResetList** 或 **wdContinueList**），该常量表明是否接续上一列表格式。

**语法**

**express.CanContinuePreviousList(ListTemplate)**

*express*   一个代表 **List** 对象的变量。

**参数**

| **名称**       | **必选/可选** | **数据类型**     | **说明**                             |
| -------------- | ------------- | ---------------- | ------------------------------------ |
| *ListTemplate* | 必选          | **ListTemplate** | 一个应用于文档中先前段落的列表模板。 |

**说明**

本方法返回指定列表格式的**“继续前一列表”**和**“重新开始编号”**选项状态，这些选项位于**“项目符号和编号”**对话框中。设置 **ApplyListTemplate** 方法的 ContinuePreviousList 参数可更改这些选项的设置。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例检查是否可以接续前面得编号，则应用当前的列表模板，其编号接续先前列表的编号。所选内容必须位于第二个列表中，否则导致出错 */ function test() { 	let selListFmt = Application.Selection.Range.ListFormat 	let canConti = selListFmt.CanContinuePreviousList(selListFmt.ListTemplate) 	if (canConti !== wdContinueDisabled ) { 		selListFmt.ApplyListTemplate(selListFmt.ListTemplate, true) 	} }` |

#### **List.ConvertNumbersToText**

更改指定的 **List** 对象中的列表编号和 LISTNUM 域。

**语法**

**express.ConvertNumbersToText()**

*express*   一个代表 **List** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例将第一个列表的编号转换为文本 */ Application.ActiveDocument.Lists.Item(1).ConvertNumbersToText()` |

#### **List.CountNumberedItems**

返回指定 **List** 对象中的项目符号或编号项目和 LISTNUM 域的数量。

**语法**

**express.CountNumberedItems()**

*express*   一个代表 **List** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例返回第一个List对象中数量 */ Application.ActiveDocument.Lists.Item(1).CountNumberedItems()` |

#### **List.RemoveNumbers**

删除指定列表中的编号或项目符号。

**语法**

**express.RemoveNumbers(NumberType)**

*express*   一个代表 **List** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型**     | **说明**           |
| ------------ | ------------- | ---------------- | ------------------ |
| *NumberType* | 可选          | **WdNumberType** | 要删除的编号类型。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 移除第一个列表List得项目编号或符号 */ Application.ActiveDocument.Lists.Item(1).RemoveNumbers() ` |

**成员属性**

#### **List.ListParagraphs**

返回一个 **ListParagraphs** 集合，该集合代表列表、文档或范围中的所有编号段落。只读。

**语法**

**express.ListParagraphs**

*express*   一个代表 **List** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 将第一个列表中得第一个段落设置为双下划线 */ Application.ActiveDocument.Lists.Item(1).ListParagraphs.Item(1).Range.Underline = wdUnderlineDouble` |

#### **List.Range**

返回一个 **Range** 对象，该对象代表指定对象所包含的文档部分。

**语法**

**express.Range**

*express*   一个代表 **List** 对象的变量。

#### **List.SingleListTemplate**

如果整个列表都使用同一列表模板，则该属性值为 **True**。只读 **Boolean** 类型。

**语法**

**express.SingleListTemplate**

*express*   一个代表 **List** 对象的变量。

适用环境：web

适用平台：windows/linux