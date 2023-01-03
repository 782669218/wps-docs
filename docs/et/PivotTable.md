**PivotTable**



代表工作表上的数据透视表。

**说明**

**PivotTable** 对象是 **PivotTables** 集合的成员。**PivotTables** 集合包含某一张工作表上的所有 **PivotTable** 对象。

因为对数据透视表进行编程可能会很复杂，所以，最方便的做法是将数据透视表操作录制到宏中，然后再修订所录制的宏代码。

使用 **PivotTables**( *index*)（其中 *index* 是数据透视表索引号或名称）可返回一个 **PivotTable** 对象。下例使 Sheet3 上第一张数据透视表中的字段“Year”成为行字段。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `Worksheets.Item("Sheet3").PivotTables(1).PivotFields("Year").Orientation = xlRowField` |

**方法**

|                                                              | 名称                  | 说明                                                         |
| ------------------------------------------------------------ | --------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddDataField**      | 将数据字段添加到数据透视表中。返回一个 **PivotField** 对象，该对象表示新的数据字段。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddFields**         | 向数据透视表或数据透视图中添加行字段、列字段和页字段。       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AllocateChanges**   | table { word-break:break-all; }对基于 OLAP 数据源的数据透视表中所有编辑过的单元格执行回写操作。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CalculatedFields**  | 返回一个 **CalculatedFields** 集合，该集合表示指定数据透视表中的所有计算字段。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ChangeConnection**  | table { word-break:break-all; }更改指定 **PivotTable** 的连接。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ChangePivotCache**  | table { word-break:break-all; }更改指定 **PivotTable** 的 **PivotCache**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ClearAllFilters**   | table { word-break:break-all; }**ClearAllFilters** 方法将删除当前应用于数据透视表的所有筛选。包括删除 **PivotTable** 对象的 **PivotFilters** 集合中的所有筛选，删除应用的所有手动筛选以及将报表筛选区域中的所有筛选字段设置为默认项。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ClearTable**        | **ClearTable** 方法用于清除数据透视表。清除数据透视表包括删除所有字段以及删除应用于数据透视表的所有筛选和排序。此方法将数据透视表重置为该表刚创建而尚未添加任何字段时的状态。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CommitChanges**     | table { word-break:break-all; }对基于 OLAP 数据源的数据透视表的数据源执行提交操作。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ConvertToFormulas** | **ConvertToFormulas** 方法是 ET 中的新方法，可用于将数据透视表转换为多维数据集公式。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CreateCubeFile**    | 创建数据透视表的多维数据集文件，该数据透视表与联机分析处理 (OLAP)?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源相连接。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **DiscardChanges**    | table { word-break:break-all; }放弃基于 OLAP 数据源的数据透视表中经过编辑的单元格中的所有更改。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **GetData**           | 返回数据透视表中数据字段的值。                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **GetPivotData**      | 返回一个 **Range** 对象，并返回有关数据透视表中数据项的信息。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ListFormulas**      | 在分离工作表上创建数据透视表的计算项和计算字段的列表。       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PivotCache**        | 返回一个 **PivotCache** 对象，该对象表示指定数据透视表的缓存。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PivotFields**       | 返回一个对象，该对象表示数据透视表中单个字段（**PivotField** 对象）或可见及隐藏字段集合（**PivotFields** 对象）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PivotSelect**       | 选定数据透视表的一部分。                                     |

**属性**

|                                                              | 名称                              | 说明                                                         |
| ------------------------------------------------------------ | --------------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**                   | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CacheIndex**                    | 返回或设置数据透视表缓存的索引号。**Long** 类型，可读写。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CalculatedMembers**             | 返回一个 **CalculatedMembers** 集合，该集合表示 OLAP 数据透视表的所有计算成员和计算方法。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CalculatedMembersInFilter**     | 返回或设置是否估算筛选器中的 OLAP 服务器计算成员。可读/写    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ChangeList**                    | 返回 **PivotTableChangeList** 集合，该集合表示已对基于 OLAP 数据源的指定数据透视表所做更改的列表。只读 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ColumnFields**                  | 返回一个对象，该对象表示单个的数据透视表字段（**PivotField** 对象）或当前显示为列字段的所有字段的集合（**PivotFields** 对象）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ColumnGrand**                   | 如果数据透视表显示列总计，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ColumnRange**                   | 返回一个 **Range** 对象，该对象表示数据透视表中包含列区域的区域。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CompactLayoutColumnHeader**     | 指定透视数据表在压缩行布局表单中时标题中显示的标题。可读写 **String**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CompactLayoutRowHeader**        | 指定透视数据表在压缩行布局表单中时行标题中显示的标题。可读写 **String**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CompactRowIndent**              | 当启用压缩行布局表单时，返回或设置透视项目的缩进增量。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                       | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CubeFields**                    | 返回 **CubeFields** 集合。每个 **CubeField** 对象都包含多维数据集?（多维数据集：一种 OLAP 数据结构。多维数据集包含维度，如“国家/地区/省（或市/自治区）/市（或县）”，还包括数据字段，如“销售额”。维度将各种类型的数据组织到带有明细数据级别的分层结构中，而数据字段度量数量。）字段元素的属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DataBodyRange**                 | 返回一个 **Range** 对象，它代表数据透视表中数值区域。只读。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DataFields**                    | 返回一个对象，该对象表示单个数据透视表字段（**PivotField** 对象）或当前显示为数据字段的所有字段集合（**PivotFields** 对象）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DataLabelRange**                | 返回一个 **Range** 对象，该对象表示在数据透视表中包含数据字段的标签的区域。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DataPivotField**                | 返回一个 **PivotField** 对象，该对象表示数据透视表中所有数据字段。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayContextTooltips**        | 控制是否显示数据透视表单元格的工具提示。可读写 **Boolean**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayEmptyColumn**            | 如果对数值轴的 OLAP 提供程序的查询中包括非空 MDX 关键字，则返回 **True**。在结果集中，OLAP 提供程序不返回空列。如果省略非空关键字，则返回 **False**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayEmptyRow**               | 如果对分类轴的 OLAP 提供程序的查询中包括非空 MDX 关键字，则返回 **True**。在结果集中，OLAP 提供程序不返回空行。如果省略非空关键字，则返回 **False**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayErrorString**            | 如果数据透视表在有错误的单元格中显示用户自定义的错误字符串，则该值为 **True**。默认值为 **False**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayFieldCaptions**          | 控制是否在网格中显示筛选器按钮以及行和列的透视字段标题。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayImmediateItems**         | 返回或设置 **Boolean**，用于指明当数据透视表的数据区域为空时，行和列区域中的项是否可见。如果该属性为 **False**，则当数据透视表的数据区域为空时，将隐藏行和列区域中的项。默认值为 **True**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayMemberPropertyTooltips** | 控制是否在工具提示中显示成员属性。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayNullString**             | 如果数据透视表在包含空值的单元格中显示用户自定义的字符串，则该值为 **True**。默认值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EnableDataValueEditing**        | 如果为 **True**，则当用户覆盖数据透视表数据区域中的值时禁用警告。设置为 **True** 也可以使用户更改先前无法更改的数据值。默认值为 **False**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EnableDrilldown**               | 如果启用“显示明细数据”，则该值为 **True**。默认值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EnableFieldDialog**             | 如果当用户双击数据透视表字段时，**“数据透视表字段”**对话框可用，则该属性值为 **True**。默认值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EnableFieldList**               | 如果为 **False**，则禁止显示数据透视表字段列表的功能。如果已经显示字段列表，则该列表将消失。默认值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EnableWizard**                  | 如果**“数据透视表向导”**可用，则该属性值为 **True**。默认值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EnableWriteback**               | 返回或设置是否对指定的数据透视表启用回写到数据源。默认值为 **False**。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ErrorString**                   | 返回或设置一个 **String** 值，它代表 **DisplayErrorString** 属性为 **True** 时如果单元格中有错误而显示的字符串。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FieldListSortAscending**        | 控制数据透视表字段列表中字段的排序顺序。当此属性设置为 **True** 时，字段按升序顺序排序。当它设置为 **False** 时，字段按数据源顺序排序。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **GrandTotalName**                | 返回或设置显示在指定数据透视表的总计列或行标题中的文本串标志。默认值为字符串“Grand Total”。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HiddenFields**                  | 返回一个对象，该对象表示单个数据透视表字段（**PivotField** 对象），或表示当前未显示为行、列、页或数据字段的所有字段的集合（**PivotFields** 对象）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **InGridDropZones**               | 此属性用于为 **PivotTable** 对象切换网格中的拖放区域。在一些情况下，它还会影响数据透视表的布局。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **InnerDetail**                   | 当最内部行或列字段的 **ShowDetail** 属性设为 **True** 时，返回或设置这些详细数据的字段名称。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LayoutRowDefault**              | 此属性指定初次将透视字段添加到数据透视表中时它们的布局设置。可读/写 **xlLayoutRowType** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Location**                      | 获取或设置一个 **String** 类型的值，该值代表指定 **PivotTable** 主体中左上角的单元格。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ManualUpdate**                  | 如果数据透视表仅在用户请求时重新计算，则该值为 **True**。默认值为 **False**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MDX**                           | 返回一个 **String** 类型的值，该值表示将发送给提供程序以填充当前数据透视表视图的多维表达式 (MDX)。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MergeLabels**                   | 如果指定的数据透视表的外部行项、列项、分类汇总和总计标志使用合并单元格，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**                          | 返回或设置一个 **String** 值，它代表对象的名称。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NullString**                    | 返回或设置当 **DisplayNullString** 为 **True** 时，在包含 null 值的单元格中显示的字符串。默认值为空字符串 ("")。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PageFieldOrder**                | 返回或设置页字段在数据透视表布局上的顺序。可为以下 **XlOrder** 常量之一：**xlDownThenOver** 或 **xlOverThenDown**。默认常量是 **xlDownThenOver**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PageFields**                    | 返回一个对象，该对象表示单个数据透视表字段（**PivotField** 对象），或者当前显示为页字段的所有字段的一个集合（**PivotFields** 对象）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PageFieldStyle**                | 返回或设置用于边界页字段区域中的样式。默认值为 null 字符串（默认时无样式）。**String** 类型，可读写。**语法** |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PageFieldWrapCount**            | 返回或设置数据透视表中每行或每列的页字段数目。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PageRange**                     | 返回一个 **Range** 对象，该对象表示包含数据透视表中页区域的区域。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PageRangeCells**                | 返回一个 **Range** 对象，该对象仅表示指定数据透视表中（包含页字段和项下拉列表）的单元格。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                        | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PivotColumnAxis**               | 返回 **PivotAxis** 对象，该对象表示整个列轴。只读 **PivotAxis** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PivotFormulas**                 | 返回一个 **PivotFormulas** 对象，该对象表示指定数据透视表的公式集合。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PivotRowAxis**                  | 返回 **PivotAxis** 对象，该对象表示整个行轴。只读 **PivotAxis** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PivotSelection**                | 以标准数据透视表的选定区域格式返回或设置数据透视表的选定区域。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PivotSelectionStandard**        | 返回或设置用英语（美国）设置的标准数据透视表?（数据透视表：一种交互的、交叉制表的 ET 报表，用于对多种来源（包括 ET 的外部数据）的数据（如数据库记录）进行汇总和分析。）格式的数据透视表选定内容的 **String** 类型的数值。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PreserveFormatting**            | 当使用数据透视、筛选或更改页字段项等操作刷新或重新计算报告时，如果格式处于被保护状态，则为 **True**。对于查询表，如果数据前五行的任何常规格式设置将应用于查询表数据的新行，则此属性为 **True**。对未使用的单元格不进行设置。如果将查询表应用的最近一次自动套用格式应用于新数据行，则此属性为 **False**。默认值是 **True**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PrintDrillIndicators**          | 指定是否随数据透视表打印深化指示符。可读写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PrintTitles**                   | 如果基于数据透视表设置工作表的打印标题，则该属性值为 **True**。如果使用工作表的打印标题，则该属性值为 **False**。默认值为 **False**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RefreshDate**                   | 返回数据透视表或缓存最近一次刷新的日期。**Date** 型，只读。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RefreshName**                   | 返回最近一次刷新数据透视表的人员的名字。**String** 型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RepeatItemsOnEachPrintedPage**  | 当打印指定的数据透视表时，如果每页第一行上都显示行、列和项标志，则该值为 **True**。如果仅在第一页上打印这些标志，则该值为 **False**。默认值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RowFields**                     | 返回一个对象，该对象表示数据透视表中的单个字段（**PivotField** 对象），或者当前未显示为行字段的所有字段的一个集合（**PivotFields** 对象）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RowGrand**                      | 如果数据透视表显示行的总数，则该属性值为 **True**。**Boolean** 类型，可读写。 |

**成员方法**

#### **PivotTable.AddDataField**

将数据字段添加到数据透视表中。返回一个 **PivotField** 对象，该对象表示新的数据字段。

**语法**

**express.AddDataField(Field, Caption, Function )**

*express*   一个代表 **PivotTable** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**                                                     |
| ---------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Field*    | 必选          | **Object**   | 服务器上的唯一字段。如果源数据是联机分析处理 (OLAP)?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。），则唯一字段是多维数据集字段。如果源数据不是 OLAP（非 OLAP 源数据?（非 OLAP 源数据：非 OLAP 源数据是数据透视表或数据透视图使用的基本数据，该数据来自 OLAP 数据库之外的源。这些数据源包括关系数据库、 ET 工作表中的清单以及文本文件数据库。）），则唯一字段是数据透视表字段。 |
| *Caption*  | 可选          | **Variant**  | 数据透视表中使用的标签，用于识别该数据字段。                 |
| *Function* | 可选          | **Variant**  | 在已添加数据字段中执行的函数。                               |

**返回值**

PivotField

**示例**

本示例将标题为“Total Score”的数据字段添加到名为“PivotTable1”的数据透视表中。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 本示例假定存在这样的表格，该表格中包含一个标题为“Score”的列。 |

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `function AddMoreFields() {     let pTab = ActiveSheet.PivotTables("PivotTable1")     pTab.**AddDataField**(ActiveSheet.PivotTables("PivotTable1").PivotFields("Score"), "Total Score") }` |

#### **PivotTable.AddFields**

向数据透视表或数据透视图中添加行字段、列字段和页字段。

**语法**

**express.AddFields(RowFields, ColumnFields, PageFields, AddToTable)**

*express*   一个代表 **PivotTable** 对象的变量。

**参数**

| **名称**       | **必选/可选** | **数据类型** | **说明**                                                     |
| -------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *RowFields*    | 可选          | **Variant**  | 指定要作为行添加或要添加到类别坐标轴中的字段名（或字段名数组）。 |
| *ColumnFields* | 可选          | **Variant**  | 指定要作为列添加或要添加到序列坐标轴中的字段名（或字段名数组）。 |
| *PageFields*   | 可选          | **Variant**  | 指定要作为页添加或要添加到页区域中的字段名（或字段名数组）。 |
| *AddToTable*   | 可选          | **Variant**  | 仅适用于数据透视表。如果为 True，则将指定的字段添加到报表中（不替换现有字段）。如果为 False，则用新的字段替换现有的字段。默认值为 False。 |

**返回值**

Variant

**说明**

必须指定其中某个字段参数。

字段名指定由 **PivotField** 对象的 **SourceName** 属性返回的唯一名称。

这个方法对 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源不可用。

**示例**

本示例以“Status”字段和“Closed_By”字段替换 Sheet1 的第一个数据透视表的现有列字段。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `Worksheets.Item("Sheet1").PivotTables(1).**AddFields**(null,["Status", "Closed_By"])` |

#### **PivotTable.AllocateChanges**

table { word-break:break-all; }

对基于 OLAP 数据源的数据透视表中所有编辑过的单元格执行回写操作。

**语法**

**express.AllocateChanges()**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

table { word-break:break-all; }

**AllocateChanges** 方法将执行 **UPDATE CUBE** 语句，以应用自上次提交应用更改操作以来，或创建数据透视表以来（如果从未执行过提交应用更改），数据透视表的值区域中发生的所有更改。如果对基于非 OLAP 数据源的数据透视表执行此方法，那么将生成运行时错误。

#### **PivotTable.CalculatedFields**

返回一个 **CalculatedFields** 集合，该集合表示指定数据透视表中的所有计算字段。只读。

**语法**

**express.CalculatedFields()**

*express*   一个代表 **PivotTable** 对象的变量。

**返回值**

CalculatedFields

**示例**

本示例使计算结果字段不能被拖至行。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `for(let i=1;i <= Worksheets.Item(1).PivotTables("Pivot1").**CalculatedFields**().Count;i++) {     Worksheets.Item(1).PivotTables("Pivot1").CalculatedFields().Item(i).DragToRow = false }` |

#### **PivotTable.ChangeConnection**

table { word-break:break-all; }

更改指定 **PivotTable** 的连接。

**语法**

**express.ChangeConnection(conn)**

*express*   一个代表 **PivotTable** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**           | **说明**                                               |
| -------- | ------------- | ---------------------- | ------------------------------------------------------ |
| *conn*   | 必选          | **WorkbookConnection** | 一个代表数据透视表的新连接的 WorkbookConnection 对象。 |

**说明**

table { word-break:break-all; }

**ChangeConnection** 方法只能用于连接到外部数据源的 **PivotTable**。如果将 **ChangeConnection** 方法与使用工作表中存储的数据作为数据源的 **PivotTable** 一起使用，将发生运行时错误。

#### **PivotTable.ChangePivotCache**

table { word-break:break-all; }

更改指定 **PivotTable** 的 **PivotCache**。

**语法**

**express.ChangePivotCache(bstr)**

*express*   一个代表 **PivotTable** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *bstr*   | 必选          | **String**   | 一个 PivotTable 或 PivotCache 对象，该对象代表指定 PivotTable 的新 PivotCache。 |

**说明**

table { word-break:break-all; }

**ChangePivotCache** 方法只能用于使用工作表中存储的数据作为数据源的 **PivotTable**。如果将 **ChangePivotCache** 方法与连接到外部数据源的 **PivotTable** 一起使用，将发生运行时错误。

#### **PivotTable.ClearAllFilters**

table { word-break:break-all; }

**ClearAllFilters** 方法将删除当前应用于数据透视表的所有筛选。包括删除 **PivotTable** 对象的 **PivotFilters** 集合中的所有筛选，删除应用的所有手动筛选以及将报表筛选区域中的所有筛选字段设置为默认项。

**语法**

**express.ClearAllFilters()**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

table { word-break:break-all; }

调用 **ClearAllFilters** 方法后，**PivotFilters** 集合将为空。

#### **PivotTable.ClearTable**

**ClearTable** 方法用于清除数据透视表。清除数据透视表包括删除所有字段以及删除应用于数据透视表的所有筛选和排序。此方法将数据透视表重置为该表刚创建而尚未添加任何字段时的状态。

**语法**

**express.ClearTable()**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

**ClearTable** 函数不使用任何参数，可用于关系数据透视表和 OLAP 数据透视表。

**示例**

以下示例清除活动工作表中的数据透视表。

| 示例代码                                  |
| ----------------------------------------- |
| `ActiveSheet.PivotTables(1).ClearTable()` |

#### **PivotTable.CommitChanges**

table { word-break:break-all; }

对基于 OLAP 数据源的数据透视表的数据源执行提交操作。

**语法**

**express.CommitChanges()**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

table { word-break:break-all; }

**CommitChanges** 方法将 **COMMIT TRANSACTION** 语句发送给 OLAP 服务器，并且清除通过输入值而编辑的所有单元格，但是不清除值单元格中的公式。如果对基于非 OLAP 数据源的数据透视表执行此方法，那么将生成运行时错误。

#### **PivotTable.ConvertToFormulas**

**ConvertToFormulas** 方法是 ET 中的新方法，可用于将数据透视表转换为多维数据集公式。可读/写 **Boolean** 类型。

**语法**

**express.ConvertToFormulas(ConvertFilters)**

*express*   一个代表 **PivotTable** 对象的变量。

**参数**

| **名称**         | **必选/可选** | **数据类型** | **说明**                                             |
| ---------------- | ------------- | ------------ | ---------------------------------------------------- |
| *ConvertFilters* | 必选          | **Boolean**  | 包含 True 或 False，以指明 ReportFilter 区域的状态。 |

**说明**

此参数控制是否转换数据透视表的 **ReportFilter** 区域。

**示例**

在以下示例中，不转换 **ReportFilter** 区域。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `function ConvertToCubeFormulas() {     ActiveSheet.PivotTables("PivotTable1").ConvertToFormulas(false) }` |

#### **PivotTable.CreateCubeFile**

创建数据透视表的多维数据集文件，该数据透视表与联机分析处理 (OLAP)?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源相连接。

**语法**

**express.CreateCubeFile(File, Measures, Levels, Members, Properties)**

*express*   一个代表 **PivotTable** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------ | ------------- | ------------ | ------------------------------------------------------------ |
| *File*       | 必选          | **String**   | 要创建的多维数据集文件的名称。如果该文件已存在，则将被覆盖。 |
| *Measures*   | 可选          | **Variant**  | 度量的唯一名称的数组，该度量是扇面的一部分。                 |
| *Levels*     | 可选          | **Variant**  | 字符串数组。每个数组项都是一个唯一的级别名称。该参数表示扇面中的层次结构的最低级别。 |
| *Members*    | 可选          | **Variant**  | 字符串数组的数组。元素按顺序对应于由 Levels 数组表示的分层结构。每个元素都是字符串数组的一个数组，由维（包括在扇面中）中的最高级别成员的唯一名称组成。 |
| *Properties* | 可选          | **Variant**  | 如果该值为 False，则扇区中不包含成员属性。默认值为 True      |

**返回值**

String

**示例**

本示例在驱动器 C:\ 下创建一个名为“CustomCubeFile”的多维数据集文件，扇面中不包括任何成员属性。由于本示例省略了 *Measures*、*Levels* 和 *Members* 参数，因此该多维数据集文件将停止匹配数据透视表的视图。本示例假定与 OLAP 数据源相连接的数据透视表位于活动工作表上。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `function UseCreateCubeFile() {     ActiveSheet.PivotTables(1).**CreateCubeFile**("C:\\CustomCubeFile",null,null,null,false) }` |

#### **PivotTable.DiscardChanges**

table { word-break:break-all; }

放弃基于 OLAP 数据源的数据透视表中经过编辑的单元格中的所有更改。

**语法**

**express.DiscardChanges()**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

table { word-break:break-all; }

对于基于 OLAP 数据源的数据透视表，此方法删除在值单元格中输入的所有值和公式，然后运行数据透视表更新操作从数据源中检索最新值。它将所有编辑过的值单元格的对应数据源值设置为 **NULL**，并且还对 OLAP 服务器执行 **ROLLBACK TRANSACTION** 语句。

如果尝试对基于非 OLAP 数据源的数据透视表执行此方法，则会生成运行时错误。

#### **PivotTable.GetData**

返回数据透视表中数据字段的值。

**语法**

**express.GetData(Name)**

*express*   一个代表 **PivotTable** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Name*   | 必选          | **String**   | 描述数据透视表中的单个单元格，使用的语法类似于 PivotSelect 方法或者计算项公式中的数据透表引用。 |

**返回值**

双精度

**示例**

本示例显示一月的苹果收入之和 (Data field = Revenue、Product = Apples、Month = January)。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `MsgBox(ActiveSheet.PivotTables(1).**GetData**("'Sum of Revenue' Apples January"))` |

#### **PivotTable.GetPivotData**

返回一个 **Range** 对象，并返回有关数据透视表中数据项的信息。

**语法**

**express.GetPivotData(DataField, Field1~14, Item1~14)**

*express*   一个代表 **PivotTable** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                         |
| ----------- | ------------- | ------------ | -------------------------------- |
| *DataField* | 可选          | **Variant**  | 包含数据透视表数据的字段的名称。 |
| *Field1~14* | 可选          | **Variant**  | 数据透视表中的列或行字段的名称。 |
| *Item1~14*  | 可选          | **Variant**  | Field中的项目的名称。            |

**返回值**

Range

**示例**

在本例中，ET 向用户返回仓库中椅子的数量。本示例假定数据透视表位于活动工作表上。此外，本示例还假设：在数据透视表中，数据字段的标题为“Quantity”，同时还存在标题为“Warehouse”的字段，并且在“Warehouse”字段中有标题为“Chairs”的数据项。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `function UseGetPivotData() {     // Get PivotData for the quantity of chairs in the warehouse.     let rngTableItem = ActiveCell.PivotTable.**GetPivotData**("Quantity", "Warehouse", "Chairs")     MsgBox("The quantity of chairs in the warehouse is: " + rngTableItem.Value()) }` |

#### **PivotTable.ListFormulas**

在分离工作表上创建数据透视表的计算项和计算字段的列表。

**语法**

**express.ListFormulas()**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

本方法对 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源无效。

**示例**

本示例为第一个工作表上的第一个数据透视表创建一个计算项和计算字段的列表。

| 示例代码                                               |
| ------------------------------------------------------ |
| `Worksheets.Item(1).PivotTables(1).**ListFormulas**()` |

#### **PivotTable.PivotCache**

返回一个 **PivotCache** 对象，该对象表示指定数据透视表的缓存。只读。

**语法**

**express.PivotCache()**

*express*   一个代表 **PivotTable** 对象的变量。

**返回值**

PivotCache

**示例**

本示例使第一张工作表上的第一张数据透视表的缓存在构造时进行优化。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `Worksheets.Item(1).PivotTables("Pivot1").**PivotCache**.OptimizeCache = true` |

#### **PivotTable.PivotFields**

返回一个对象，该对象表示数据透视表中单个字段（**PivotField** 对象）或可见及隐藏字段集合（**PivotFields** 对象）。只读。

**语法**

**express.PivotFields(Index)**

*express*   一个代表 **PivotTable** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                   |
| -------- | ------------- | ------------ | -------------------------- |
| *Index*  | 可选          | **Variant**  | 要返回的字段的名称或编号。 |

**返回值**

Object

**说明**

对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，不存在隐藏字段，返回的对象或集合反映了当前可见的内容。

**示例**

本示例向新工作表列表添加数据透视表的字段名称。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `let nwSheet = Worksheets.Add() nwSheet.Activate() let pvtTable = Worksheets.Item("Sheet2").Range("A1").PivotTable let rw = 0 for(let i=1;i <= pvtTable.**PivotFields**().Count;i++) {     rw++     nwSheet.Cells.Item(rw, 1).Value2 = pvtTable.PivotFields(i).Name }` |

#### **PivotTable.PivotSelect**

选定数据透视表的一部分。

**语法**

**express.PivotSelect(Name, Mode, UseStandardName)**

*express*   一个代表 **PivotTable** 对象的变量。

**参数**

| **名称**          | **必选/可选** | **数据类型**          | **说明**                                      |
| ----------------- | ------------- | --------------------- | --------------------------------------------- |
| *Name*            | 必选          | **String**            | 要选择的数据透视表的一部分。                  |
| *Mode*            | 可选          | **XlPTSelectionMode** | 指定结构化选择模式。                          |
| *UseStandardName* | 可选          | **Variant**           | 如果为 True，则表示将在其他位置运行录制的宏。 |

**说明**

只能使用指定模式来选择数据透视表中的相应项。例如，无法通过 **xlButton** 模式来选择数据和标志；同样，也无法通过 **xlDataOnly** 模式来选择按钮。

**示例**

本示例选定第一张工作表上的第一张数据透视表中的所有日期标志。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `Worksheets.Item(1).PivotTables(1).**PivotSelect**("date[All]", xlLabelOnly)` |

**成员属性**

#### **PivotTable.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **PivotTable** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test(){ let myObject = Application.ActiveWorkbook if(myObject.Application.Value == "ET") {     alert("This is an ET Application object.") } else {     alert("This is not an ET Application object.") } }` |

#### **PivotTable.CacheIndex**

返回或设置数据透视表缓存的索引号。**Long** 类型，可读写。

**语法**

**express.CacheIndex**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

如果设置 **CacheIndex** 属性，以便于数据透视表使用第二张数据透视表的高速缓存，则第一个报表的字段必须是第二个报表中字段的有效子集。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将数据透视表“Pivot1”的缓存设置成数据透视表“Pivot2”的缓存。*/ Application.Worksheets.Item(1).PivotTables("Pivot1").CacheIndex = Worksheets.Item(1).PivotTables("Pivot2").CacheIndex    ` |

#### **PivotTable.CalculatedMembers**

返回一个 **CalculatedMembers** 集合，该集合表示 OLAP 数据透视表的所有计算成员和计算方法。

**语法**

**express.CalculatedMembers**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

本属性用于联机分析处理 (OLAP)?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 源，非 OLAP 源将返回一个运行时错误。

**示例**

本示例向数据透视表添加一个设置。假定数据透视表位于一个与 OLAP 数据源相连的活动工作表上，此 OLAP 数据源中包含一个标题为“[Product].[All Products]”的字段。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   let pvtTable = Application.ActiveSheet.PivotTables(1)     // Add the calculated member.     pvtTable.CalculatedMembers.Add("[Beef]", "'{[Product].[All Products].Children}'", null, xlCalculatedSet) }` |

#### **PivotTable.CalculatedMembersInFilter**

返回或设置是否估算筛选器中的 OLAP 服务器计算成员。可读/写

**语法**

**express.CalculatedMembersInFilter**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

如果在筛选器中估算计算成员，则为 **True**；否则为 **False**。

此属性的值对应于基于 OLAP 数据源的数据透视表的**“数据透视表选项”**对话框**“汇总和筛选”**选项卡上的**“估算筛选器中的 OLAP 服务器计算成员”**复选框的设置。

#### **PivotTable.ChangeList**

返回 **PivotTableChangeList** 集合，该集合表示已对基于 OLAP 数据源的指定数据透视表所做更改的列表。只读

**语法**

**express.ChangeList**

*express*   一个代表 **PivotTable** 对象的变量。

#### **PivotTable.ColumnFields**

返回一个对象，该对象表示单个的数据透视表字段（**PivotField** 对象）或当前显示为列字段的所有字段的集合（**PivotFields** 对象）。只读。

**语法**

**express.ColumnFields**

*express*   一个代表 **PivotTable** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将数据透视表的列字段名称添加到一张新工作表上的数据清单中。*/ function test(){ let nwSheet = Worksheets.Add() nwSheet.Activate() let pvtTable = Worksheets.Item("Sheet2").Range("A1").PivotTable let rw = 0 for(let i=1;i <= pvtTable.ColumnFields().Count;i++) {     rw++     nwSheet.Cells.Item(rw, 1).Value2 = pvtTable.ColumnFields(i).Name } }` |

#### **PivotTable.ColumnGrand**

如果数据透视表显示列总计，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ColumnGrand**

*express*   一个代表 **PivotTable** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置数据透视表显示列总计。*/ function test(){ let pvtTable = Worksheets.Item("Sheet1").Range("A3").PivotTable pvtTable.ColumnGrand = true }` |

#### **PivotTable.ColumnRange**

返回一个 **Range** 对象，该对象表示数据透视表中包含列区域的区域。只读。

**语法**

**express.ColumnRange**

*express*   一个代表 **PivotTable** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为数据透视表选择列标题。*/ function test(){ Worksheets.Item("Sheet1").Activate() Range("A3").Select() ActiveCell.PivotTable.ColumnRange.Select() }` |

#### **PivotTable.CompactLayoutColumnHeader**

指定透视数据表在压缩行布局表单中时标题中显示的标题。可读写 **String**。

**语法**

**express.CompactLayoutColumnHeader**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

用户在网格中自定义这些标题时，这些标题也会反映到这些属性中。

#### **PivotTable.CompactLayoutRowHeader**

指定透视数据表在压缩行布局表单中时行标题中显示的标题。可读写 **String**。

**语法**

**express.CompactLayoutRowHeader**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

用户在网格中自定义这些标题时，这些标题也会反映到这些属性中。

#### **PivotTable.CompactRowIndent**

当启用压缩行布局表单时，返回或设置透视项目的缩进增量。可读写。

**语法**

**express.CompactRowIndent**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

默认值为 1。此设置的有效值为 0 到 ET 中指定的最大缩进。

#### **PivotTable.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **PivotTable.CubeFields**

返回 **CubeFields** 集合。每个 **CubeField** 对象都包含多维数据集?（多维数据集：一种 OLAP 数据结构。多维数据集包含维度，如“国家/地区/省（或市/自治区）/市（或县）”，还包括数据字段，如“销售额”。维度将各种类型的数据组织到带有明细数据级别的分层结构中，而数据字段度量数量。）字段元素的属性。只读。

**语法**

**express.CubeFields**

*express*   一个代表 **PivotTable** 对象的变量。

**示例**

本示例为工作表 Sheet1 上第一个基于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 的数据透视表中的数据字段，创建一个多维数据集?（多维数据集：一种 OLAP 数据结构。多维数据集包含维度，如“国家/地区/省（或市/自治区）/市（或县）”，还包括数据字段，如“销售额”。维度将各种类型的数据组织到带有明细数据级别的分层结构中，而数据字段度量数量。）字段名称列表。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){ let objNewSheet = Worksheets.Add() objNewSheet.Activate() let intRow = 1 for(let i=1;i <= Worksheets.Item("Sheet1").PivotTables(1).CubeFields.Count;i++) {     if(Worksheets.Item("Sheet1").PivotTables(1).CubeFields.Item(i).Orientation == xlDataField) {         objNewSheet.Cells.Item(intRow, 1).Value2 = Worksheets.Item("Sheet1").PivotTables(1).CubeFields.Item(i).Name         intRow++     } } }` |

#### **PivotTable.DataBodyRange**

返回一个 **Range** 对象，它代表数据透视表中数值区域。只读。

**语法**

**express.DataBodyRange**

*express*   一个代表 **PivotTable** 对象的变量。

#### **PivotTable.DataFields**

返回一个对象，该对象表示单个数据透视表字段（**PivotField** 对象）或当前显示为数据字段的所有字段集合（**PivotFields** 对象）。只读。

**语法**

**express.DataFields**

*express*   一个代表 **PivotTable** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向一张新工作表中添加数据透视表数据字段的名称列表。*/ function test(){ let nwSheet = Worksheets.Add() nwSheet.Activate() let pvtTable = Worksheets.Item("Sheet2").Range("A1").PivotTable let rw = 0 for(let i=1;i <= pvtTable.DataFields().Count;i++) {     rw++     nwSheet.Cells.Item(rw, 1).Value2 = pvtTable.DataFields(i).Name } }` |

#### **PivotTable.DataLabelRange**

返回一个 **Range** 对象，该对象表示在数据透视表中包含数据字段的标签的区域。只读。

**语法**

**express.DataLabelRange**

*express*   一个代表 **PivotTable** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例选定数据透视表中的数据字段标签。*/ function test(){ Worksheets.Item("Sheet1").Activate() Range("A3").Select() ActiveCell.PivotTable.DataLabelRange.Select() }` |

#### **PivotTable.DataPivotField**

返回一个 **PivotField** 对象，该对象表示数据透视表中所有数据字段。只读。

**语法**

**express.DataPivotField**

*express*   一个代表 **PivotTable** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将第二个 PivotItem 对象移到第一个位置。假定数据透视表位于活动的工作表上，并且数据透视表包含数据字段。*/ function test(){  let pvtTable = ActiveSheet.PivotTables(1)     // Move second PivotItem to the first position in PivotTable.     pvtTable.DataPivotField.PivotItems(2).Position = 1 }` |

#### **PivotTable.DisplayContextTooltips**

控制是否显示数据透视表单元格的工具提示。可读写 **Boolean**。

**语法**

**express.DisplayContextTooltips**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

如果 **DisplayContextTooltips** 设置为 **False**，将不显示数据透视表单元格的上下文工具提示。如果此属性设置为 **True**，将显示行、列和数据区中数据透视表单元格的工具提示。

#### **PivotTable.DisplayEmptyColumn**

如果对数值轴的 OLAP 提供程序的查询中包括非空 MDX 关键字，则返回 **True**。在结果集中，OLAP 提供程序不返回空列。如果省略非空关键字，则返回 **False**。**Boolean** 类型，可读写。

**语法**

**express.DisplayEmptyColumn**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

数据透视表必须与一个联机分析处理 (OLAP)?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源相连，以避免运行时错误。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例确定数据透视表中空列的显示设置。本示例假定与 OLAP 数据源相连的数据透视表位于活动工作表上。*/ function test(){  let pvtTable = ActiveSheet.PivotTables(1)     // Determine display setting for empty columns.     if(pvtTable.DisplayEmptyColumn == false) {         alert("Empty columns will not be displayed.")     }     else {         alert("Empty columns will be displayed.")     } }` |

#### **PivotTable.DisplayEmptyRow**

如果对分类轴的 OLAP 提供程序的查询中包括非空 MDX 关键字，则返回 **True**。在结果集中，OLAP 提供程序不返回空行。如果省略非空关键字，则返回 **False**。**Boolean** 类型，可读写。

**语法**

**express.DisplayEmptyRow**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

数据透视表必须与一个联机分析处理 (OLAP)?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源相连，以避免运行时错误。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例确定用于数据透视表中的空行的显示设置。本示例假定与 OLAP 数据源相连的数据透视表位于活动工作表上。*/ function test(){   let pvtTable = ActiveSheet.PivotTables(1)     // Determine display setting for empty rows.     if(pvtTable.DisplayEmptyRow == false) {         alert("Empty rows will not be displayed.")     }     else {         alert("Empty rows will be displayed.")     } }` |

#### **PivotTable.DisplayErrorString**

如果数据透视表在有错误的单元格中显示用户自定义的错误字符串，则该值为 **True**。默认值为 **False**。**Boolean** 类型，可读写。

**语法**

**express.DisplayErrorString**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

使用 **ErrorString** 属性可设置用户自定义的错误字符串。

当透视处理计算字段时，该属性对于消除被零除错误尤其有用。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例使数据透视表在有错误的单元格中显示连字符。*/ function test(){ Worksheets.Item(1).PivotTables("Pivot1").ErrorString = "-" Worksheets.Item(1).PivotTables("Pivot1").DisplayErrorString = true }` |

#### **PivotTable.DisplayFieldCaptions**

控制是否在网格中显示筛选器按钮以及行和列的透视字段标题。可读/写。

**语法**

**express.DisplayFieldCaptions**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

默认值为 **True**。

#### **PivotTable.DisplayImmediateItems**

返回或设置 **Boolean**，用于指明当数据透视表的数据区域为空时，行和列区域中的项是否可见。如果该属性为 **False**，则当数据透视表的数据区域为空时，将隐藏行和列区域中的项。默认值为 **True**。

**语法**

**express.DisplayImmediateItems**

*express*   一个代表 **PivotTable** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例确定如何创建数据透视表并通知用户。本示例假定数据透视表位于活动工作表上。*/ function test(){ let pvtTable = ActiveSheet.PivotTables(1) // Determine how the PivotTable was created. if(pvtTable.DisplayImmediateItems == true) { alert("Fields have been added to the row or column areas for the PivotTable report.") } else { alert("The PivotTable was created by using object-model calls.") } }` |

#### **PivotTable.DisplayMemberPropertyTooltips**

控制是否在工具提示中显示成员属性。可读/写 **Boolean** 类型。

**语法**

**express.DisplayMemberPropertyTooltips**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

如果 **DisplayMemberPropertyTooltips** 设置为 **False**，则将不会在工具提示中显示任何成员属性。如果它设置为 **True**，则将在工具提示中显示成员属性。

#### **PivotTable.DisplayNullString**

如果数据透视表在包含空值的单元格中显示用户自定义的字符串，则该值为 **True**。默认值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.DisplayNullString**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

使用 **NullString** 属性可设置自定义空字符串。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例使数据透视表在包含空值的单元格中显示“NA”。*/ function test(){ Application.Worksheets.Item(1).PivotTables("Pivot1").NullString = "NA" Application.Worksheets.Item(1).PivotTables("Pivot1").DisplayNullString = true }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例使数据透视表在包含空值的单元格中显示零 (0)。*/ Application.Worksheets.Item(1).PivotTables("Pivot1").DisplayNullString = false` |

#### **PivotTable.EnableDataValueEditing**

如果为 **True**，则当用户覆盖数据透视表数据区域中的值时禁用警告。设置为 **True** 也可以使用户更改先前无法更改的数据值。默认值为 **False**。**Boolean** 类型，可读写。

**语法**

**express.EnableDataValueEditing**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

在数据值上进行的任何编辑，在刷新时都将丢失。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例确定对覆盖数据区域中的值时的警告设置，并通知用户。本示例假定数据透视表位于活动工作表上。*/ function test(){  let pvtTable = ActiveSheet.PivotTables(1)     // Determine alert setting.     if(pvtTable.EnableDataValueEditing == false) {         alert("Alert is enabled.")     }     else {         alert("Alert is disabled.")     } }` |

#### **PivotTable.EnableDrilldown**

如果启用“显示明细数据”，则该值为 **True**。默认值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.EnableDrilldown**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

为数据透视表设置该属性将对该表中的所有字段有效。

对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，该属性值总是 **True**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对第一张工作表上第一个数据透视表中的所有字段禁用明细。*/ Application.Worksheets.Item(1).PivotTables("Pivot1").EnableDrilldown = false ` |

#### **PivotTable.EnableFieldDialog**

如果当用户双击数据透视表字段时，**“数据透视表字段”**对话框可用，则该属性值为 **True**。默认值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.EnableFieldDialog**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

为数据透视表设置该属性将对该表中的所有字段有效。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对“Year”字段禁用“数据透视表字段”对话框。*/ Application.Worksheets.Item(1).PivotTables("Pivot1").PivotFields("Year").EnableFieldDialog = false    ` |

#### **PivotTable.EnableFieldList**

如果为 **False**，则禁止显示数据透视表字段列表的功能。如果已经显示字段列表，则该列表将消失。默认值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.EnableFieldList**

*express*   一个代表 **PivotTable** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例确定是否能显示数据透视表的字段列表并通知用户。本示例假定数据透视表位于活动工作表上。*/ function test(){  let pvtTable = ActiveSheet.PivotTables(1)     // Determine if field list can be displayed.     if(pvtTable.EnableFieldList == true) {         alert("The field list for the PivotTable can be displayed.")     }     else {         alert("The field list for the PivotTable cannot be displayed.")     }  }` |

#### **PivotTable.EnableWizard**

如果**“数据透视表向导”**可用，则该属性值为 **True**。默认值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.EnableWizard**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

在设置该属性时，字段井并不会显示在工作表上。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对第一张工作表上第一个数据透视表禁用“数据透视表向导”.*/ Application.Worksheets.Item(1).PivotTables("Pivot1").EnableWizard = false` |

#### **PivotTable.EnableWriteback**

返回或设置是否对指定的数据透视表启用回写到数据源。默认值为 **False**。可读/写。

**语法**

**express.EnableWriteback**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

对于 OLAP 数据源，如果将 **EnableWriteback** 属性设置为 **True**，那么将启用回写，并禁用用户在覆盖数据透视表数据区中的值时出现的警告。对于非 OLAP 数据源，如果将 **EnableWriteback** 属性设置为 **True**，则在代码中启用回写，还将允许用户更改以前无法更改的数据值。

不能将 **PivotTable** 对象的 **EnableWriteback** 和 **EnableDataValueEditing** 属性同时设置为 **True**。

如果将 **EnableDataValueEditing** 属性设置为 **True**，然后再将 **EnableWriteback** 属性设置为 **True**，**EnableDataValueEditing** 属性就会自动设置为 **False**，数据透视表将被刷新并且对数据值进行的所有编辑操作都将丢失。

如果将 **EnableWriteback** 属性设置为 **True**，然后再将 **EnableDataValueEditing** 属性设置为 **True**，**EnableWriteback** 属性就会自动设置为 **False**，不刷新数据透视表，但会还原数据源值。

对于非 OLAP 数据源，设置此属性将会生成运行时错误。

#### **PivotTable.ErrorString**

返回或设置一个 **String** 值，它代表 **DisplayErrorString** 属性为 **True** 时如果单元格中有错误而显示的字符串。

**语法**

**express.ErrorString**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

该属性的默认值是个空字符串 ("")。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*对于指定数据透视表中有错误的单元格，本示例将在单元格中显示连字符。*/ function test(){ Application.Worksheets.Item(1).PivotTables("Pivot1").ErrorString = "-" Application.Worksheets.Item(1).PivotTables("Pivot1").DisplayErrorString = true }` |

#### **PivotTable.FieldListSortAscending**

控制数据透视表字段列表中字段的排序顺序。当此属性设置为 **True** 时，字段按升序顺序排序。当它设置为 **False** 时，字段按数据源顺序排序。可读/写。

**语法**

**express.FieldListSortAscending**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

不适用于与 OLAP 数据源相连的数据透视表。

#### **PivotTable.GrandTotalName**

返回或设置显示在指定数据透视表的总计列或行标题中的文本串标志。默认值为字符串“Grand Total”。**String** 类型，可读写。

**语法**

**express.GrandTotalName**

*express*   一个代表 **PivotTable** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动工作表的第二个数据透视表中总计标题的标志设置为“Regional Total”。*/ Application.ActiveSheet.PivotTables("PivotTable2").GrandTotalName = "Regional Total"    ` |

#### **PivotTable.HiddenFields**

返回一个对象，该对象表示单个数据透视表字段（**PivotField** 对象），或表示当前未显示为行、列、页或数据字段的所有字段的集合（**PivotFields** 对象）。只读。

**语法**

**express.HiddenFields**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，该属性总返回一个空集合。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将指定隐藏字段的名称添加进一张新工作表中的列表中。*/ function test(){ let nwSheet = Worksheets.Add() nwSheet.Activate() let pvtTable = Worksheets.Item("Sheet2").Range("A1").PivotTable let rw = 0 for(let i=1;i <= pvtTable.HiddenFields.Count;i++) {     rw++     nwSheet.Cells.Item(rw, 1).Value2 = pvtTable.HiddenFields.Item(i).Name } }` |

#### **PivotTable.InGridDropZones**

此属性用于为 **PivotTable** 对象切换网格中的拖放区域。在一些情况下，它还会影响数据透视表的布局。可读/写 **Boolean** 类型。

**语法**

**express.InGridDropZones**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

当 **InGridDropZones** 属性设置为 **True** 时，存在网格中的拖放区域。当此属性设置为 **False** 时，不存在网格中的拖放区域。数据透视表的布局也会随此属性一起改变。

#### **PivotTable.InnerDetail**

当最内部行或列字段的 **ShowDetail** 属性设为 **True** 时，返回或设置这些详细数据的字段名称。**String** 类型，可读写。

**语法**

**express.InnerDetail**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

该属性不可用于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例当最内部行或列字段的 ShowDetail 属性设为 True 时，显示这些详细数据的字段名称。*/ function test(){ let pvtTable = Application.Worksheets.Item("Sheet1").Range("A3").PivotTable alert(pvtTable.InnerDetail) }` |

#### **PivotTable.LayoutRowDefault**

此属性指定初次将透视字段添加到数据透视表中时它们的布局设置。可读/写 **xlLayoutRowType** 类型。

**语法**

**express.LayoutRowDefault**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

**xlLayoutRowType** 可以是 **xlCompactRow**、**xlTabularRow** 或 **xlOutlineRow**。

#### **PivotTable.Location**

获取或设置一个 **String** 类型的值，该值代表指定 **PivotTable** 主体中左上角的单元格。可读/写。

**语法**

**express.Location**

*express*   一个代表 **PivotTable** 对象的变量。

#### **PivotTable.ManualUpdate**

如果数据透视表仅在用户请求时重新计算，则该值为 **True**。默认值为 **False**。**Boolean** 类型，可读写。

**语法**

**express.ManualUpdate**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

程序终止之后，以及在 Microsoft Visual Basic 编辑器的立即窗口执行了该语句之后，本属性的值立即设为 **False**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例使数据透视表仅在用户请求时进行重新计算。*/ function test(){ Application.Worksheets.Item(1).PivotTables("Pivot1").ManualUpdate = true }` |

#### **PivotTable.MDX**

返回一个 **String** 类型的值，该值表示将发送给提供程序以填充当前数据透视表视图的多维表达式 (MDX)。只读。

**语法**

**express.MDX**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

查询非联机分析处理 (OLAP) 数据透视表的此值，或者当不存在数据透视表视图（无数据项）时，将返回一个运行错误。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回数据透视表的 MDX 字符串。本示例假定数据透视表位于活动工作表上。*/ function test(){  let pvtTable = ActiveSheet.PivotTables(1)     alert("The MDX string for the PivotTable is: " + pvtTable.MDX) }` |

#### **PivotTable.MergeLabels**

如果指定的数据透视表的外部行项、列项、分类汇总和总计标志使用合并单元格，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.MergeLabels**

*express*   一个代表 **PivotTable** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例使第一张工作表上的第一个数据透视表使用合并单元格外部行项、列项、分类汇总和总计的标志。*/ function test(){ Application.Worksheets.Item(1).PivotTables(1).MergeLabels = true }` |

#### **PivotTable.Name**

返回或设置一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

*express*   一个代表 **PivotTable** 对象的变量。

#### **PivotTable.NullString**

返回或设置当 **DisplayNullString** 为 **True** 时，在包含 null 值的单元格中显示的字符串。默认值为空字符串 ("")。**String** 类型，可读写。

**语法**

**express.NullString**

*express*   一个代表 **PivotTable** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例使数据透视表在包含空值的单元格中显示“NA”。*/ function test(){ Application.Worksheets.Item(1).PivotTables("Pivot1").NullString = "NA" Application.Worksheets.Item(1).PivotTables("Pivot1").DisplayNullString = true }` |

#### **PivotTable.PageFieldOrder**

返回或设置页字段在数据透视表布局上的顺序。可为以下 **XlOrder** 常量之一：**xlDownThenOver** 或 **xlOverThenDown**。默认常量是 **xlDownThenOver**。**Long** 类型，可读写。

**语法**

**express.PageFieldOrder**

*express*   一个代表 **PivotTable** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例使数据透视表先在一行中显示三个页字段，然后再显示下一个新行。*/ function test(){ Application.Worksheets.Item(1).PivotTables("Pivot1").PageFieldOrder = xlOverThenDown Application.Worksheets.Item(1).PivotTables("Pivot1").PageFieldWrapCount = 3 }` |

#### **PivotTable.PageFields**

返回一个对象，该对象表示单个数据透视表字段（**PivotField** 对象），或者当前显示为页字段的所有字段的一个集合（**PivotFields** 对象）。只读。

**语法**

**express.PageFields**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

一个层次只能包含一个页字段。

对基于数据透视表高速缓存的数据透视表而言，返回的数据透视表字段的集合反应了当前高速缓存中的内容。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向一张新工作表上的列表中添加页字段名称。*/ function test(){ let nwSheet = Application.Worksheets.Add() nwSheet.Activate() let pvtTable = Worksheets.Item("Sheet2").Range("A1").PivotTable let rw = 0 for(let i=1;i <= pvtTable.PageFields().Count;i++) {     rw++     nwSheet.Cells.Item(rw, 1).Value2 = pvtTable.PageFields(i).Name } }` |

#### **PivotTable.PageFieldStyle**

返回或设置用于边界页字段区域中的样式。默认值为 null 字符串（默认时无样式）。**String** 类型，可读写。

**语法**

**语法**

**express.PageFieldStyle**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

此样式用作背景区域的默认样式，并且在用户指定格式前起作用。当对某字段进行数据透视时，从页字段区域移到其他位置的单元格保留此样式。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将第一张工作表上的第一个数据透视表中的页字段区域设置为 PurpleAndGold 样式。*/ Application.Worksheets.Item(1).PivotTables("Pivot1").PageFieldStyle = "PurpleAndGold"` |

#### **PivotTable.PageFieldWrapCount**

返回或设置数据透视表中每行或每列的页字段数目。**Long** 类型，可读写。

**语法**

**express.PageFieldWrapCount**

*express*   一个代表 **PivotTable** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例使数据透视表先在一行中显示三个页字段，然后再显示下一个新行。*/ function test(){ let pTab = Application.Worksheets.Item(1).PivotTables("Pivot1") pTab.PageFieldOrder = xlOverThenDown pTab.PageFieldWrapCount = 3 }` |

#### **PivotTable.PageRange**

返回一个 **Range** 对象，该对象表示包含数据透视表中页区域的区域。只读。

**语法**

**express.PageRange**

*express*   一个代表 **PivotTable** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例选择数据透视表中的页面页眉。*/ function test(){ Application.Worksheets.Item("Sheet1").Activate() Range("A3").Select() Application.ActiveCell.PivotTable.PageRange.Select() }` |

#### **PivotTable.PageRangeCells**

返回一个 **Range** 对象，该对象仅表示指定数据透视表中（包含页字段和项下拉列表）的单元格。

**语法**

**express.PageRangeCells**

*express*   一个代表 **PivotTable** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例仅选择包含页字段和数据项下拉列表的数据透视表中的单元格。*/ Application.Worksheets.Item(1).PivotTables(1).PageRangeCells.Select()` |

#### **PivotTable.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **PivotTable** 对象的变量。

#### **PivotTable.PivotColumnAxis**

返回 **PivotAxis** 对象，该对象表示整个列轴。只读 **PivotAxis** 类型。

**语法**

**express.PivotColumnAxis**

*express*   一个代表 **PivotTable** 对象的变量。

#### **PivotTable.PivotFormulas**

返回一个 **PivotFormulas** 对象，该对象表示指定数据透视表的公式集合。只读。

**语法**

**express.PivotFormulas**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，该属性将返回一个空集合。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){ let r = 0 for(let i=1;i <= Application.ActiveSheet.PivotTables(1).PivotFormulas.Count;i++) {     r++     Cells.Item(r, 1).Value2 = Application.ActiveSheet.PivotTables(1).PivotFormulas.Item(i).Formula } }` |

#### **PivotTable.PivotRowAxis**

返回 **PivotAxis** 对象，该对象表示整个行轴。只读 **PivotAxis** 类型。

**语法**

**express.PivotRowAxis**

*express*   一个代表 **PivotTable** 对象的变量。

#### **PivotTable.PivotSelection**

以标准数据透视表的选定区域格式返回或设置数据透视表的选定区域。**String** 类型，可读写。

**语法**

**express.PivotSelection**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

设置该属性等同于在 *Mode* 参数设置为 **xlDataAndLabel** 时调用 **PivotSelect** 方法。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在第一张工作表上的第一张数据透视表中选择名为“Bob”的销售员的数据和标签。*/ Application.Worksheets.Item(1).PivotTables(1).PivotSelection = "Salesman[Bob]"` |

#### **PivotTable.PivotSelectionStandard**

返回或设置用英语（美国）设置的标准数据透视表?（数据透视表：一种交互的、交叉制表的 ET 报表，用于对多种来源（包括 ET 的外部数据）的数据（如数据库记录）进行汇总和分析。）格式的数据透视表选定内容的 **String** 类型的数值。可读写。

**语法**

**express.PivotSelectionStandard**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

**PivotSelection** 属性是国际通用的，而 **PivotSelectionStandard** 方法则不是。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例选择数据透视表中标题为“1.57”的字段，并在其前面插入一个空白列字段。本示例假定数据透视表位于包含标题为“1.57”的列字段的活动工作表上。*/ function test(){  let pvtTable = Application.ActiveSheet.PivotTables(1)     pvtTable.PivotSelectionStandard = "1.57"     Selection.Insert() }` |

#### **PivotTable.PreserveFormatting**

当使用数据透视、筛选或更改页字段项等操作刷新或重新计算报告时，如果格式处于被保护状态，则为 **True**。对于查询表，如果数据前五行的任何常规格式设置将应用于查询表数据的新行，则此属性为 **True**。对未使用的单元格不进行设置。如果将查询表应用的最近一次自动套用格式应用于新数据行，则此属性为 **False**。默认值是 **True**。

**语法**

**express.PreserveFormatting**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

对于数据库查询表，默认的格式设置为 **xlSimple**。

刷新查询表时，将对查询表应用新的自动套用格式样式。只要 **PreserveFormatting** 的值为 **False**，则 AutoFormat（自动套用格式）就会被设置为 **None**。这样，任何在 **PreserveFormatting** 被设置为 **False** 或在查询表刷新之前设置的自动套用格式都不会起作用，且相应产生的查询表也不会被应用任何格式。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例保留第一张工作表上的第一个数据透视表的格式。*/ Application.Worksheets.Item(1).PivotTables("Pivot1").PreserveFormatting = true ` |

此示例演示了将 **PreserveFormatting** 设置为 **False** 后，如何操作才能将“自动套用格式”设置为 **xlRangeAutoFormatNone**，而并不是指定的 **xlRangeAutoFormatColor1** 格式。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){ let QTab = Application.Workbooks.Item(1).Worksheets.Item(1).QueryTables.Item(1) QTab.Range.AutoFormat = xlRangeAutoFormatColor1 QTab.PreserveFormatting = false QTab.Refresh() }` |

#### **PivotTable.PrintDrillIndicators**

指定是否随数据透视表打印深化指示符。可读写 **Boolean** 类型。

**语法**

**express.PrintDrillIndicators**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

如果随数据透视表打印深化指示符，则返回 **True**；如果不随数据透视表打印深化指示符，则返回 **False**。

当 **ShowDrillIndicators** 属性设置为 **False** 时，**PrintDrillIndicators** 属性没有效果。如果在数据透视表中未显示深化指示符，则不会打印这些指示符。

#### **PivotTable.PrintTitles**

如果基于数据透视表设置工作表的打印标题，则该属性值为 **True**。如果使用工作表的打印标题，则该属性值为 **False**。默认值为 **False**。**Boolean** 类型，可读写。

**语法**

**express.PrintTitles**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

行打印标题设置为包含数据透视表的列字段项的行。列打印标题设置为包含行项目的列。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例指定当打印活动工作表上的第四个数据透视表时，同时也打印该工作表的打印标题集。*/ Application.ActiveSheet.PivotTables("PivotTable4").PrintTitles = true` |

#### **PivotTable.RefreshDate**

返回数据透视表或缓存最近一次刷新的日期。**Date** 型，只读。

**语法**

**express.RefreshDate**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，在每次查询后都会更新此属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例显示数据透视表最近一次刷新的日期。*/ function test(){ Set pvtTable = Application.Worksheets("Sheet1").Range("A3").PivotTable dateString = Format(pvtTable.RefreshDate, "Long Date") MsgBox "The data was last refreshed on " & dateString }` |

#### **PivotTable.RefreshName**

返回最近一次刷新数据透视表的人员的名字。**String** 型，只读。

**语法**

**express.RefreshName**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，在每次查询后都会更新此属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例显示最近一次更新数据透视表的用户的名字。*/ function test(){ let pvtTable = Worksheets.Item("Sheet1").Range("A3").PivotTable alert("The data was last refreshed by " + pvtTable.RefreshName) }` |

#### **PivotTable.RepeatItemsOnEachPrintedPage**

当打印指定的数据透视表时，如果每页第一行上都显示行、列和项标志，则该值为 **True**。如果仅在第一页上打印这些标志，则该值为 **False**。默认值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.RepeatItemsOnEachPrintedPage**

*express*   一个代表 **PivotTable** 对象的变量。

**说明**

对于工作表，ET 打印行和列标签，而不是打印标题。使用 **PrintTitles** 属性可判断是否为数据透视表设置了打印标题。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*在打印活动工作表的第四个数据透视表时，本示例使 ET 在每页上重复打印标签。*/ Application.ActiveSheet.PivotTables("PivotTable4").RepeatItemsOnEachPrintedPage = true` |

#### **PivotTable.RowFields**

返回一个对象，该对象表示数据透视表中的单个字段（**PivotField** 对象），或者当前未显示为行字段的所有字段的一个集合（**PivotFields** 对象）。只读。

**语法**

**express.RowFields**

*express*   一个代表 **PivotTable** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在新工作表上的清单中添加数据透视表的行字段名称。*/ function test(){ let nwSheet = Application.Worksheets.Add() nwSheet.Activate() let pvtTable =  Worksheets("Sheet2").Range("A1").PivotTable let rw = 0 for(let i =   1;  i < = pvtTable.RowFields().Count; i++) {     rw = rw + 1     nwSheet.Cells.Item(rw, 1).Value2 = pvtTable.RowFields(i).Name } }` |

#### **PivotTable.RowGrand**

如果数据透视表显示行的总数，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.RowGrand**

*express*   一个代表 **PivotTable** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置数据透视表以显示行的总计。*/ function test(){ let pvtTable = Application.Worksheets.Item("Sheet1").Range("A3").PivotTable pvtTable.RowGrand = true }` |

适用环境：web

适用平台：windows/linux