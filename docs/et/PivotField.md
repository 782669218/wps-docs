**PivotField**



代表数据透视表中的一个字段。

**说明**

**PivotField** 对象是 **PivotFields** 集合的成员。**PivotFields** 集合包含数据透视表中的所有字段，包括隐藏字段。

下列属性返回数据透视表字段的子集，在某些情况下，使用这些属性会更为方便：

- **ColumnFields** 属性
- **DataFields** 属性
- **HiddenFields** 属性
- **PageFields** 属性
- **RowFields** 属性
- **VisibleFields** 属性

使用 **PivotFields**( *index*)（其中 *index* 是字段名称或索引号）可返回一个 **PivotField** 对象。下例使 Sheet3 上第一张数据透视表中的字段“Year”成为行字段。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Worksheets.Item("sheet3").PivotTables(1).PivotFields("year").Orientation = xlRowField` |

**方法**

|                                                              | 名称                  | 说明                                                         |
| ------------------------------------------------------------ | --------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddPageItem**       | 向具有多个项的页面字段添加其他项。                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AutoShow**          | 显示指定数据透视表中行字段、页字段或列字段顶部或底部数据项的个数。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AutoSort**          | 为数据透视表建立自动字段排序规则。                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CalculatedItems**   | 返回一个 **CalculatedItems** 集合，该集合表示指定数据透视表中的所有计算项。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ClearAllFilters**   | 调用此方法将删除当前应用于透视字段的所有筛选。包括从透视字段的 **PivotFilters** 集合中删除所有筛选以及删除应用于透视字段的所有手动筛选。如果透视字段位于报表筛选区域中，则所选项将被设置为默认项。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ClearLabelFilters** | 此方法将删除透视字段的 **PivotFilters** 集合中的所有标签筛选或所有日期筛选。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ClearManualFilter** | 提供一个简便方法，可将数据透视表中透视字段的所有项的 **Visible** 属性设置为 **True**，并清空 OLAP 数据透视表中的 **HiddenItemsList** 和 **VisibleItemsList** 集合。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ClearValueFilters** | 调用此方法将删除透视字段的 **PivotFilters** 集合中的所有值筛选。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**            | 删除对象。                                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **DrillTo**           | **DrillTo** 方法支持从另一个透视字段深化到指定的透视字段。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PivotItems**        | 返回一个对象，该对象表示指定字段中的单个数据透视表项（**PivotItem** 对象）或所有可见项和隐藏项的集合（**PivotItems** 对象）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **RepeatAllLabels**   | 指定是否要对指定数据透视表中的所有透视字段重复项目标签。     |

**属性**

|                                                              | 名称                           | 说明                                                         |
| ------------------------------------------------------------ | ------------------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AllItemsVisible**            | 用于检索指示是否对透视字段应用任何手动筛选的 Boolean 值。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**                | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoShowCount**              | 返回指定数据透视表字段中自动显示的首项号或末项号。**Long** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoShowField**              | 返回数据字段的名称，该字段用于判断在指定数据透视表字段中自动显示的是首项还是末项。**String** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoShowRange**              | 如果指定数据透视表字段中自动显示首项，则返回 **xlTop**；如果自动显示末项，则返回 **xlBottom**。**Long** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoShowType**               | 如果启用指定数据透视表字段的 **AutoShow**，则返回 **xlAutomatic**；如果禁用 **AutoShow**，则返回 **xlManual**。**Long** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoSortCustomSubtotal**     | 返回用于对指定数据透视表字段进行自动排序的自定义分类汇总的名称。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoSortField**              | 返回用于对指定数据透视表字段进行自动排序的数据字段的名称。**String** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoSortOrder**              | 返回用于对指定数据透视表字段进行自动排序的次序。可为 **XlSortOrder** 常量之一。**Long** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoSortPivotLine**          | table { word-break:break-all; }返回用于对指定数据透视表字段进行自动排序的 PivotLine 的名称。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BaseField**                  | 返回或设置自定义计算的基准字段。本属性仅对数据字段有效。**Variant** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BaseItem**                   | 返回或设置自定义计算基准字段的数据项，仅对数据字段有效。**Variant** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Calculation**                | 返回或设置一个 **XlPivotFieldCalculation** 值，它代表指定字段执行的计算类型。此属性仅对数据字段有效。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Caption**                    | table { word-break:break-all; }返回一个 **String** 值，它代表数据透视字段的标签文本。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ChildField**                 | 返回一个 **PivotField** 对象，该对象表示特定字段的子字段（如果该字段已分组且有子字段）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ChildItems**                 | 返回一个对象，它代表一个数据透视表数据项（**PivotItem** 对象）或所有数据项的集合（**PivotItems** 对象），该集合可以由指定字段中的子项组成，也可以由指定数据项中的子项组成。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                    | table { word-break:break-all; }返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CubeField**                  | 返回派生出指定数据透视表字段的 **CubeField** 对象。只读。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CurrentPage**                | 返回或设置页字段的当前页显示（仅对页字段有效）。**PivotItem** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CurrentPageList**            | 返回或设置对应于项目列表的字符串数组，该项目列表包含于数据透视表的多项目页字段中。**Variant** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CurrentPageName**            | 返回或设置指定数据透视表上的当前显示页。该页名称将出现在页字段中。注意，只有当已存在当前显示页时，本属性才有效。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DatabaseSort**               | 如果为 **True**，则允许手动更改数据透视表字段中项目的位置。如果该字段中没有手动定位的项目，则返回 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DataRange**                  | 返回一个 **Range** 对象，如下表所示。只读。                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DataType**                   | 返回一个 **XlPivotFieldDataType** 值，它代表数据透视表字段中的数据类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayAsCaption**           | table { word-break:break-all; }此属性用于将透视字段的成员属性显示为标题。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayAsTooltip**           | table { word-break:break-all; }此属性用于指定在工具提示中是否显示透视字段的特定成员属性。可读写。**Boolean**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayInReport**            | table { word-break:break-all; }此属性用于指定在数据透视表中是否显示指定的成员属性 PivotField。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DragToColumn**               | 如果指定字段能被拖动到列位置上，则为 **True**。默认值是 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DragToData**                 | 如果指定字段可被拖动到数据位置上，则为 **True**。默认值是 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DragToHide**                 | 如果通过将字段拖离数据透视表可隐藏该字段，则为 **True**。默认值是 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DragToPage**                 | 如果字段可被拖动到页位置上，则为 **True**。默认值是 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DragToRow**                  | 如果字段可被拖动到行位置上，则为 **True**。默认值是 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DrilledDown**                | 如果指定数据透视表字段或数据透视表项的标志设置为“drilled”（展开或可见），则为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EnableItemSelection**        | 如果为 **False**，则在用户界面中禁止使用下拉字段的功能。默认值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EnableMultiplePageItems**    | table { word-break:break-all; }用于指定对于页面区域中的字段是否在筛选器下拉列表中显示复选框。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Formula**                    | table { word-break:break-all; }返回或设置一个 **String** 值，它代表 A1 样式表示法和宏语言中的对象的公式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Function**                   | 返回或设置对数据透视表字段汇总时所使用的函数（仅用于数据字段）。**XlConsolidationFunction** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **GroupLevel**                 | 返回一组字段中指定字段的位置（如果该字段是分组字段集合中的成员）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Hidden**                     | table { word-break:break-all; }此属性用于隐藏 OLAP 层次结构的各个级别。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HiddenItems**                | 返回一个对象，该对象表示单个隐藏数据透视表项（**PivotItem** 对象），或表示指定字段中所有隐藏项的集合（**PivotItems** 对象）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HiddenItemsList**            | 返回或设置一个字符串数组，该数组为数据透视表字段的隐藏项。**Variant** 类型，可读写 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IncludeNewItemsInFilter**    | table { word-break:break-all; }在将手动筛选应用于透视字段时，此属性允许开发人员指定是应跟踪排除的项目还是应跟踪包含的项目。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IsCalculated**               | 如果数据透视表字段为计算字段或计算项，则此属性为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IsMemberProperty**           | 如果数据透视表字段包含成员属性，则返回 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LabelRange**                 | 返回一个 **Range** 对象，它代表包含字段标签的单元格。只读。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LayoutBlankLine**            | 如果在数据透视表的指定行字段后插入了一个空行，则该值为 **True**。默认值为 **False**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LayoutCompactRow**           | table { word-break:break-all; }指定在选择行时是否压缩透视字段（在一列中显示多个透视字段的项目）。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LayoutForm**                 | 返回或设置指定的数据透视表项出现的方式，即以表格格式还是以分级显示格式显示。**XlLayoutFormType** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LayoutPageBreak**            |                                                              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LayoutSubtotalLocation**     | 返回或设置与指定字段相关（在其上面或下面）的数据透视表字段分类汇总的位置。**XlSubtototalLocationType** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MemberPropertyCaption**      | table { word-break:break-all; }设置 **MemberPropertyCaption** 属性可控制将哪个成员属性用作指定级别的标题。可读写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MemoryUsed**                 | 返回对象当前使用的内存大小，以字节表示。**Long** 型，只读。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**                       | table { word-break:break-all; }返回或设置一个 **String** 值，它代表对象的名称。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NumberFormat**               | table { word-break:break-all; }返回或设置一个 **String** 值，它代表对象的格式代码。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Orientation**                | 返回或设置一个 **XlPivotFieldOrientation** 值，它代表字段在指定的数据透视表中的位置。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                     | table { word-break:break-all; }返回指定对象的父对象。只读。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ParentField**                | 返回一个 **PivotField** 对象，该对象表示作为指定对象的组父级的数据透视表字段。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ParentItems**                | 返回一个对象，该对象表示指定字段中的单个数据透视表项（**PivotItem** 对象），或者作为组父级的所有项的一个集合（**PivotItems** 对象）。指定字段必须是另一个字段的组父级。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PivotFilters**               | table { word-break:break-all; }返回或设置指定 **PivotField** 对象的 PivotFilters。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Position**                   | table { word-break:break-all; }返回一个 Variant 值，他代表数据透视表上所有字段的位置，包括属于所有行、列、页以及数据方向上的字段（第一字段、第二字段、第三字段等等）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PropertyOrder**              | 只对属于成员属性字段的数据透视表字段有效。返回一个 **Long** 类型的数值，该数值表示成员属性在其所属的多维数据集字段内的显示位置。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PropertyParentField**        | 返回一个 **PivotField** 对象，该对象表示字段中属性所从属的字段。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RepeatLabels**               | table { word-break:break-all; }返回或设置在数据透视表中是否对指定的透视字段重复项目标签。可读/写 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ServerBased**                | 如果指定数据透视表的数据源为外部数据源，并且只检索与选定页字段相匹配的数据项，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowAllItems**               | 如果显示数据透视表中的所有项目（即使这些项目中不包含汇总数据），则该值为 **True**。默认值为 **False**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowDetail**                 | table { word-break:break-all; }获取或设置指定的 **PivotField** 是否显示明细数据。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowingInAxis**              | table { word-break:break-all; }指明透视字段当前在数据透视表中是否可见。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SourceCaption**              | table { word-break:break-all; }**SourceCaption** 属性只适用于 OLAP 数据透视表，并从 OLAP 服务器中返回透视字段的原始标题。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SourceName**                 | table { word-break:break-all; }返回一个 **String** 值，它代表指定的对象出现在指定的数据透视表的原始源数据中时的名称。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **StandardFormula**            | 返回或设置一个 **String** 值，该值指定使用标准英语（美国）格式的公式。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SubtotalName**               | 返回或设置显示在指定数据透视表的分类汇总列或行标题中的文本字符串标志。默认值为“Subtotal”。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Subtotals**                  | 返回或设置与指定字段同时显示的分类汇总。仅对非数据字段有效。**Variant** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TotalLevels**                | 返回当前字段组中的字段总数。如果字段没有分组或数据源基于的是 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。），则 **TotalLevels** 将返回值 1。**Long** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **UseMemberPropertyAsCaption** | table { word-break:break-all; }此属性用于控制是否将成员属性标题用于透视字段的 PivotItem 标题。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Value**                      | table { word-break:break-all; }返回或设置一个 **String** 值，它代表数据透视表中指定的字段的名称。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **VisibleItems**               | 返回一个对象，该对象表示指定字段中的单个可见数据透视表项（**PivotItem** 对象），或指定字段中的所有可见项的集合（**PivotItems** 对象）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **VisibleItemsList**           | 返回或设置一个 **Variant** 类型的值，该值指定一个字符串数组，字符串代表应用于透视字段的手动筛选中的包含项。可读/写。 |

**成员方法**

#### **PivotField.AddPageItem**

向具有多个项的页面字段添加其他项。

**语法**

**express.AddPageItem(Item, ClearList)**

*express*   一个代表 **PivotField** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Item*      | 必选          | **String**   | PivotItem 对象的源名称，该名称与特定的联机分析处理 (OLAP)?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 成员的唯一名称相对应。 |
| *ClearList* | 可选          | **Variant**  | 如果为 False（默认值），则向现有列表中添加一个页面项。如果为 True，则删除所有当前项，然后添加 Item。 |

**说明**

为了避免运行时错误，数据源必须是 OLAP 源，选择的字段当前必须位于页面位置中，并且 **EnableMultiplePageItems** 属性必须设置为 **True**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*在本例中，ET 添加一个源名称为“[Product].[All Products].[Food].[Eggs]”的页面项。本示例假定 OLAP 数据透视表位于活动工作表上。*/ function test(){  // The source is an OLAP database and you can manually reorder items.     ActiveSheet.PivotTables(1).CubeFields.Item("[Product]"). EnableMultiplePageItems = true 								     // Add the page item titled "[Product].[All Products].[Food].[Eggs]".     ActiveSheet.PivotTables(1).PivotFields("[Product]").AddPageItem ("[Product].[All Products].[Food].[Eggs]") }` |

#### **PivotField.AutoShow**

显示指定数据透视表中行字段、页字段或列字段顶部或底部数据项的个数。

**语法**

**express.AutoShow(Type, Range, Count, Field)**

*express*   一个代表 **PivotField** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Type*   | 必选          | **Long**     | 使用 xlAutomatic 可使指定数据透视表显示与指定条件匹配的项。使用 xlManual 可禁用此功能。 |
| *Range*  | 必选          | **Long**     | 所要显示的项的起始位置。可以是下列两个常量之一：xlTop 或 xlBottom。 |
| *Count*  | 必选          | **Long**     | 要显示的项数。                                               |
| *Field*  | 必选          | **String**   | 基础数据字段的名称。必须指定唯一名称（由 SourceName 属性返回）而不是显示的名称。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例按销售总额大小只显示销售总额最高的两个公司：*/ Application.ActiveSheet.PivotTables("Pivot1").PivotFields("Company").AutoShow(xlAutomatic, xlTop, 2, "Sum of Sales")` |

#### **PivotField.AutoSort**

为数据透视表建立自动字段排序规则。

**语法**

**express.AutoSort(Order, Field, PivotLine, CustomSubtotal)**

*express*   一个代表 **PivotField** 对象的变量。

**参数**

| **名称**         | **必选/可选** | **数据类型** | **说明**                                                     |
| ---------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Order*          | 必选          | **Long**     | 指定排序顺序的 XlSortOrder 的常量之一。                      |
| *Field*          | 必选          | **String**   | 排序关键字段的名称。必须指定唯一名称（由 SourceName 属性返回的名称），而不是显示的名称。 |
| *PivotLine*      | 可选          | **Variant**  | 列上的一行或数据透视表中的一行。                             |
| *CustomSubtotal* | 可选          | **Variant**  | 自定义分类汇总字段。                                         |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本例以销售总额为基准对“Company”字段按降序排序。*/ Application.ActiveSheet.PivotTables(1).PivotFields("Company").AutoSort(xlDescending, "Sum of Sales")    ` |

#### **PivotField.CalculatedItems**

返回一个 **CalculatedItems** 集合，该集合表示指定数据透视表中的所有计算项。只读。

**语法**

**express.CalculatedItems()**

*express*   一个代表 **PivotField** 对象的变量。

**返回值**

CalculatedItems

**说明**

对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，该方法返回一个长度为零的集合。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例创建计算项及其公式的列表。*/ function test(){ let pt = Application.Worksheets.Item(1).PivotTables(1) let r = 0 for (let i = 1; i <= pt.PivotFields("Sales").CalculatedItems().Count; i++){     r++     Application.Worksheets.Item(3).Cells.Item(r, 1).Value2 = pt.PivotFields("Sales").CalculatedItems().Item(i).Name     Application.Worksheets.Item(3).Cells.Item(r, 2).Value2 = pt.PivotFields("Sales").CalculatedItems().Item(i).Formula } }` |

#### **PivotField.ClearAllFilters**

调用此方法将删除当前应用于透视字段的所有筛选。包括从透视字段的 **PivotFilters** 集合中删除所有筛选以及删除应用于透视字段的所有手动筛选。如果透视字段位于报表筛选区域中，则所选项将被设置为默认项。

**语法**

**express.ClearAllFilters()**

*express*   一个代表 **PivotField** 对象的变量。

#### **PivotField.ClearLabelFilters**

此方法将删除透视字段的 **PivotFilters** 集合中的所有标签筛选或所有日期筛选。

**语法**

**express.ClearLabelFilters()**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

下表列出了此方法将删除的各种标签筛选类型。

|                                     |
| ----------------------------------- |
| **xlCaptionEquals**                 |
| **xlCaptionDoesNotEqual**           |
| **xlCaptionIsGreaterThan**          |
| **xlCaptionIsGreaterThanOrEqualTo** |
| **xlCaptionIsLessThan**             |
| **xlCaptionIsLessThanOrEqualTo**    |
| **xlCaptionBeginsWith**             |
| **xlCaptionDoesNotBeginWith**       |
| **xlCaptionEndsWith**               |
| **xlCaptionDoesNotEndWith**         |
| **xlCaptionContains**               |
| **xlCaptionDoesNotContain**         |
| **xlCaptionIsBetween**              |
| **xlCaptionIsNotBetween**           |
|                                     |
|                                     |

下表列出了此方法将删除的各种日期筛选类型。

|                                 |
| ------------------------------- |
| **xlSpecificDate**              |
| **xlNotSpecificDate**           |
| **xlBefore**                    |
| **xlBeforeOrEqualTo**           |
| **xlAfter**                     |
| **xlAfterOrEqualTo**            |
| **xlDateBetween**               |
| **xlDateNotBetween**            |
| **xlDateToday**                 |
| **xlDateYesterday**             |
| **xlDateTomorrow**              |
| **xlDateThisWeek**              |
| **xlDateLastWeek**              |
| **xlDateNextWeek**              |
| **xlDateThisMonth**             |
| **xlDateLastMonth**             |
| **xlDateNextMonth**             |
| **xlDateThisQuarter**           |
| **xlDateLastQuarter**           |
| **xlDateNextQuarter**           |
| **xlDateThisYear**              |
| **xlDateLastYear**              |
| **xlDateNextYear**              |
| **xlYearToDate**                |
| **xlAllDatesInPeriodQuarter1**  |
| **xlAllDatesInPeriodQuarter2**  |
| **xlAllDatesInPeriodQuarter3**  |
| **xlAllDatesInPeriodQuarter4**  |
| **xlAllDatesInPeriodJanuary**   |
| **xlAllDatesInPeriodFebruary**  |
| **xlAllDatesInPeriodMarch**     |
| **xlAllDatesInPeriodApril**     |
| **xlAllDatesInPeriodMay**       |
| **xlAllDatesInPeriodJune**      |
| **xlAllDatesInPeriodJuly**      |
| **xlAllDatesInPeriodAugust**    |
| **xlAllDatesInPeriodSeptember** |
| **xlAllDatesInPeriodOctober**   |
| **xlAllDatesInPeriodNovember**  |
| **xlAllDatesInPeriodDecember**  |

#### **PivotField.ClearManualFilter**

提供一个简便方法，可将数据透视表中透视字段的所有项的 **Visible** 属性设置为 **True**，并清空 OLAP 数据透视表中的 **HiddenItemsList** 和 **VisibleItemsList** 集合。

**语法**

**express.ClearManualFilter()**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

此方法适用于数据透视表中的 **PivotField** 对象和 OLAP 数据透视表中的 **CubeField** 对象。如果对 OLAP 数据透视表中的透视字段调用此方法，将返回运行时错误。

调用此方法后，以下集合将为空：**HiddenItemsList**、**HiddenItems**、**VisibleItemsList** 和 **VisibleItems**。

#### **PivotField.ClearValueFilters**

调用此方法将删除透视字段的 **PivotFilters** 集合中的所有值筛选。

**语法**

**express.ClearValueFilters()**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

下表列出了此方法将删除的各种值筛选类型。

|                                   |
| --------------------------------- |
| **xlTopCount**                    |
| **xlBottomCount**                 |
| **xlTopPercent**                  |
| **xlBottomPercent**               |
| **xlTopSum**                      |
| **xlBottomSum**                   |
| **xlValueEquals**                 |
| **xlValueDoesNotEqual**           |
| **xlValueIsGreaterThan**          |
| **xlValueIsGreaterThanOrEqualTo** |
| **xlValueIsLessThan**             |
| **xlValueIsLessThanOrEqualTo**    |
| **xlValueIsBetween**              |
| **xlValueIsNotBetween**           |

#### **PivotField.Delete**

删除对象。

**语法**

**express.Delete()**

*express*   一个代表 **PivotField** 对象的变量。

#### **PivotField.DrillTo**

**DrillTo** 方法支持从另一个透视字段深化到指定的透视字段。

**语法**

**express.DrillTo(PivotFieldName)**

*express*   一个代表 **PivotField** 对象的变量。

**参数**

| **名称**         | **必选/可选** | **数据类型** | **说明**                       |
| ---------------- | ------------- | ------------ | ------------------------------ |
| *PivotFieldName* | 必选          | **String**   | 要深化到的 PivotField 的名称。 |

**说明**

只能在实际位于数据透视表上的字段上执行此操作。因此，要么包含所请求的透视字段的层次结构必须在数据透视表的行或列中，要么属性/关系字段必须在数据透视表的行或列中，并且放在至少一个其他属性/关系字段的旁边。而且，要深化到的字段必须在同一个层次结构或属性链中。如果不符合这些条件，则会发生运行时错误。

#### **PivotField.PivotItems**

返回一个对象，该对象表示指定字段中的单个数据透视表项（**PivotItem** 对象）或所有可见项和隐藏项的集合（**PivotItems** 对象）。只读。

**语法**

**express.PivotItems(Index)**

*express*   一个代表 **PivotField** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                 |
| -------- | ------------- | ------------ | ------------------------ |
| *Index*  | 可选          | **Variant**  | 要返回的项的名称或编号。 |

**返回值**

Variant

**说明**

对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，该集合是根据唯一的名称（由 **SourceName** 属性返回的名称），而不是根据显示的名称建立索引。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将“product”字段中所有项的名称添加到新工作表的列表中。*/ function test(){ let nwSheet = Application.Worksheets.Add()     nwSheet.Activate()     let pvtTable = Worksheets.Item("Sheet2").Range("A3").PivotTable     let rw = 0     for(let i = 1; i <= pvtTable.PivotFields("product").PivotItems.Count; i++){         rw++         nwSheet.Cells.Item(rw, 1).Value2 = pvtTable.PivotFields("product").PivotItems(i).Name     } }` |

#### **PivotField.RepeatAllLabels**

指定是否要对指定数据透视表中的所有透视字段重复项目标签。

**语法**

**express.RepeatAllLabels(Repeat)**

*express*   一个代表 **PivotField** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**                 | **说明** |
| -------- | ------------- | ---------------------------- | -------- |
| *Repeat* | 必选          | **XlPivotFieldRepeatLabels** |          |

**说明**

使用 **RepeatAllLabels** 方法对应于**“数据透视表工具设计”**选项卡的**“报表布局”**下拉列表中的**“重复所有项目标签”**和**“不重复项目标签”**命令。

若要指定是否要对单个数据透视表重复项目标签，请使用 **RepeatLabels** 属性。

**成员属性**

#### **PivotField.AllItemsVisible**

用于检索指示是否对透视字段应用任何手动筛选的 Boolean 值。只读。

**语法**

**express.AllItemsVisible**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

此属性提供了一种简单方式，便于检查是否对透视字段或多维数据集字段应用手动筛选。

对于 OLAP 数据透视表，此属性仅适用于 **CubeField** 对象。如果尝试从 OLAP 数据透视表中的 **PivotField** 对象获取此属性或设置此属性，将返回运行时错误。

对于数据透视表，此属性适用于 **PivotField** 对象。

默认值为 **True**。未应用手动筛选（无论 **IncludeNewItemsInFilter** 属性为 **True** 还是为 **False**）时，此属性将自动设置为 **True**。应用任何手动筛选（无论 **IncludeNewItemsInFilter** 属性为 **True** 还是为 **False**）时，此属性将自动设置为 **False**。

此属性直接反映透视字段或多维数据集字段的筛选下拉列表中**“全选”**复选框的状态。

#### **PivotField.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **PivotField** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() { 	let myObject = Application.ActiveWorkbook 	if (myObject.Application.Value == "ET") { 		alert("This is an ET Application object.") 	} else { 		alert("This is not an ET Application object.") 	} }` |

#### **PivotField.AutoShowCount**

返回指定数据透视表字段中自动显示的首项号或末项号。**Long** 类型，只读。

**语法**

**express.AutoShowCount**

*express*   一个代表 **PivotField** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在消息框中显示“Salesman”字段的 AutoShow 参数值。*/ function test(){ let pi = Application.Worksheets.Item(1).PivotTables(1).PivotFields("salesman") if(pi.AutoShowType == xlAutomatic){     let rn     let r = pi.AutoShowRange     if(r == xlTop){         rn = "top"     }     else{         rn = "bottom"     }     alert("PivotTable report is showing " + rn + " " + pi.AutoShowCount + " items in " + pi.Name + " field by " + pi.AutoShowField) } else{     MsgBox("PivotTable report is not using AutoShow for this field") } }` |

#### **PivotField.AutoShowField**

返回数据字段的名称，该字段用于判断在指定数据透视表字段中自动显示的是首项还是末项。**String** 类型，只读。

**语法**

**express.AutoShowField**

*express*   一个代表 **PivotField** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在消息框中显示“Salesman”字段的 AutoShow 参数值。*/ function test(){ let pi = Worksheets.Item(1).PivotTables(1).PivotFields("salesman")								 if(pi.AutoShowType == xlAutomatic){     let rn     let r = pi.AutoShowRange     if(r == xlTop){         rn = "top"     }     else{         rn = "bottom"     }     alert("PivotTable report is showing " + rn + " " + pi.AutoShowCount + " items in " + pi.Name + " field by " + pi.AutoShowField) } else{     alert("PivotTable report is not using AutoShow for this field") } }` |

#### **PivotField.AutoShowRange**

如果指定数据透视表字段中自动显示首项，则返回 **xlTop**；如果自动显示末项，则返回 **xlBottom**。**Long** 类型，只读。

**语法**

**express.AutoShowRange**

*express*   一个代表 **PivotField** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在消息框中显示“Salesman”字段的 AutoShow 参数值。*/ function test(){ let pi = Worksheets.Item(1).PivotTables(1).PivotFields("salesman") if(pi.AutoShowType == xlAutomatic){     let rn     let r = pi.AutoShowRange     if(r == xlTop){         rn = "top"     }     else{         rn = "bottom"     }     MsgBox("PivotTable report is showing " + rn + " " + pi.AutoShowCount + " items in " + pi.Name + " field by " + pi.AutoShowField) } else{     MsgBox("PivotTable report is not using AutoShow for this field") } }` |

#### **PivotField.AutoShowType**

如果启用指定数据透视表字段的 **AutoShow**，则返回 **xlAutomatic**；如果禁用 **AutoShow**，则返回 **xlManual**。**Long** 类型，只读。

**语法**

**express.AutoShowType**

*express*   一个代表 **PivotField** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在消息框中显示“Salesman”字段的 AutoShow 参数值。*/ function test(){ let pi = Application.Worksheets.Item(1).PivotTables(1).PivotFields("salesman") if(pi.AutoShowType == xlAutomatic){     let rn     let r = pi.AutoShowRange     if(r == xlTop){         rn = "top"     }     else{         rn = "bottom"     }     alert("PivotTable report is showing " + rn + " " + pi.AutoShowCount + " items in " + pi.Name + " field by " + pi.AutoShowField) } else{     alert("PivotTable report is not using AutoShow for this field") } }` |

#### **PivotField.AutoSortCustomSubtotal**

返回用于对指定数据透视表字段进行自动排序的自定义分类汇总的名称。只读。

**语法**

**express.AutoSortCustomSubtotal**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

默认值是 1（自动）。如果 **AutoSortCustomSubtotal** 属性设置是 1（自动），则数据按常规分类汇总进行排序。**AutoSortCustomSubtotal** 属性可以具有下表中列出的其中一个索引值。

| 1    | 自动         |
| ---- | ------------ |
| 2    | 求和         |
| 3    | 计数         |
| 4    | 平均值       |
| 5    | 最大值       |
| 6    | 最小值       |
| 7    | 乘积         |
| 8    | 数值计数     |
| 9    | 标准偏差     |
| 10   | 总体标准偏差 |
| 11   | 方差         |
| 12   | 总体方差     |

只有实际显示在数据透视表中的自定义分类汇总才支持排序，因此如果尝试将 **AutoSortCustomSubtotal** 设置为某个值，该值代表一个不在数据透视表视图中的自定义分类汇总，则会返回一个运行时错误。

如果基于自定义分类汇总应用排序，且从数据透视表中删除该分类汇总，则 **AutoSortCustomSubtotal** 属性将自动设置为默认值 (1)。

#### **PivotField.AutoSortField**

返回用于对指定数据透视表字段进行自动排序的数据字段的名称。**String** 类型，只读。

**语法**

**express.AutoSortField**

*express*   一个代表 **PivotField** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `本示例在消息框中显示“Product”字段的 AutoSort 参数值。  示例代码  let pi = Worksheets.Item(1).PivotTables(1).PivotFields("product") let aso switch(pi.AutoSortOrder){     case xlManual:         aso = "manual"         break;     case xlAscending:         aso = "ascending"         break;     case xlDescending:         aso = "descending"         break; } MsgBox(" sorted in " + aso + " order by " + pi.AutoSortField)` |

#### **PivotField.AutoSortOrder**

返回用于对指定数据透视表字段进行自动排序的次序。可为 **XlSortOrder** 常量之一。**Long** 类型，只读。

**语法**

**express.AutoSortOrder**

*express*   一个代表 **PivotField** 对象的变量。

**示例**

本示例在消息框中显示“Product”字段的 AutoSort 参数值。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `let pi = Worksheets.Item(1).PivotTables(1).PivotFields("product") let aso switch(pi.AutoSortOrder){     case xlManual:         aso = "manual"         break;     case xlAscending:         aso = "ascending"         break;     case xlDescending:         aso = "descending"         break; } MsgBox(" sorted in " + aso + " order by " + pi.AutoSortField)` |

#### **PivotField.AutoSortPivotLine**

table { word-break:break-all; }

返回用于对指定数据透视表字段进行自动排序的 PivotLine 的名称。只读。

**语法**

**express.AutoSortPivotLine**

*express*   一个代表 **PivotField** 对象的变量。

#### **PivotField.BaseField**

返回或设置自定义计算的基准字段。本属性仅对数据字段有效。**Variant** 类型，可读写。

**语法**

**express.BaseField**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

该属性对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源无效。

**示例**

本示例将工作表 Sheet1 上数据透视表中的数据字段设置为计算与基准字段的差，设置基准字段为“ORDER_DATE”字段，并将基准数据项设为“89-5-16”。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `let pi =  Worksheets("Sheet1").Range("A3").PivotField     pi.Calculation = xlDifferenceFrom     pi.BaseField = "ORDER_DATE"     pi.BaseItem = "5/16/89"` |

#### **PivotField.BaseItem**

返回或设置自定义计算基准字段的数据项，仅对数据字段有效。**Variant** 类型，可读写。

**语法**

**express.BaseItem**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

该属性对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源无效。

**示例**

本示例将工作表 Sheet1 上数据透视表中的数据字段设置为计算与基准字段的差，设置基准字段为“ORDER_DATE”字段，并将基准数据项设为“89-5-16”。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `let pi = Worksheets("Sheet1").Range("A3").PivotField     pi.Calculation = xlDifferenceFrom     pi.BaseField = "ORDER_DATE"     pi.BaseItem = "5/16/89"` |

#### **PivotField.Calculation**

返回或设置一个 **XlPivotFieldCalculation** 值，它代表指定字段执行的计算类型。此属性仅对数据字段有效。

**语法**

**express.Calculation**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，此属性仅能返回或被设置为 **xlNormal**。

**示例**

本示例将工作表 Sheet1 上数据透视表中的数据字段设置为计算与基准字段的差，设置基准字段为“ORDER_DATE”字段，并将基准数据项设为“89-5-16”。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `let pi = Worksheets("Sheet1").Range("A3").PivotField     pi.Calculation = xlDifferenceFrom     pi.BaseField = "ORDER_DATE"     pi.BaseItem = "5/16/89"` |

#### **PivotField.Caption**

table { word-break:break-all; }

返回一个 **String** 值，它代表数据透视字段的标签文本。

**语法**

**express.Caption**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

table { word-break:break-all; }

下表显示 **Caption** 属性及其相关属性的示例值，假设存在一个具有唯一名称为“[Europe].[France].[Paris]”的 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，并且存在一个包含名为“Paris”项的非 OLAP 数据源。

| 属性           | 值（OLAP 数据源）                 | 值（非 OLAP 数据源）        |
| -------------- | --------------------------------- | --------------------------- |
| **Caption**    | Paris                             | Paris                       |
| **Name**       | [Europe].[France].[Paris]（只读） | Paris                       |
| **SourceName** | [Europe].[France].[Paris]（只读） | （与 SQL 属性值相同；只读） |
| **Value**      | [Europe].[France].[Paris]（只读） | Paris                       |

在将一个索引值指定到 **PivotItems** 集合中时，您可以使用下表所示的语法。

| 语法（OLAP 数据源）                                | 语法（非 OLAP 数据源）         |
| -------------------------------------------------- | ------------------------------ |
| expression.PivotItems("[Europe].[France].[Paris]") | expression.PivotItems("Paris") |

在使用 **Item** 属性引用集合中的特定成员时，可以使用下表所示的文本索引名称。

| 名称（OLAP 数据源）       | 名称（非 OLAP 数据源） |
| ------------------------- | ---------------------- |
| [Europe].[France].[Paris] | Paris                  |

#### **PivotField.ChildField**

返回一个 **PivotField** 对象，该对象表示特定字段的子字段（如果该字段已分组且有子字段）。只读。

**语法**

**express.ChildField**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

如果指定字段没有子字段，则使用该属性会出现错误。

该属性对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源无效。

**示例**

本示例显示“REGION2”字段的子字段的名称。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `let pvtTable = Worksheets.Item("Sheet1").Range("A3").PivotTable MsgBox("The name of the child field is " + pvtTable.PivotFields("REGION2").ChildField.Name)` |

#### **PivotField.ChildItems**

返回一个对象，它代表一个数据透视表数据项（**PivotItem** 对象）或所有数据项的集合（**PivotItems** 对象），该集合可以由指定字段中的子项组成，也可以由指定数据项中的子项组成。只读。

**语法**

**express.ChildItems**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

此属性对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源无效。

**示例**

此示例将“vegetables”数据项的所有子数据项添加到一张新工作表上的数据清单中。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `let nwSheet = Worksheets.Add()     nwSheet.Activate() let pvtTable = Worksheets.Item("Sheet2").Range("A1").PivotTable let rw = 0 let pvtItem = pvtTable.PivotFields("product").PivotItems("vegetables").ChildItems for(let i = 1; i <= pvtItem.Count; i++){     rw = rw + 1     nwSheet.Cells.Item(rw, 1).Value2 = pvtItem(i).Name }` |

#### **PivotField.Creator**

table { word-break:break-all; }

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

table { word-break:break-all; }

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **PivotField.CubeField**

返回派生出指定数据透视表字段的 **CubeField** 对象。只读。

**语法**

**express.CubeField**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

**示例**

本示例为第一张工作表中第一个基于联机分析处理 (OLAP)?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 的数据透视表中所有分层字段创建一个多维数据集域名称列表。本示例假定在第一张工作表中存在数据透视表。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `function UseCubeField(){     let objNewSheet = Worksheets.Add()         objNewSheet.Activate()     let intRow = 1     let objPF = Worksheets.Item(1).PivotTables(1).PivotFields     for(let i = 1; i <= objPF.Count; i++){         if(objPF(i).CubeField.CubeFieldType == xlHierarchy){             objNewSheet.Cells.Item(intRow, 1).Value2 = objPF(i).Name             intRow++         }     } }` |

#### **PivotField.CurrentPage**

返回或设置页字段的当前页显示（仅对页字段有效）。**PivotItem** 类型，可读写。

**语法**

**express.CurrentPage**

*express*   一个代表 **PivotField** 对象的变量。

**示例**

本示例将 Sheet1 上数据透视表中的当前页名称返回到字符串变量 `strPgName` 中。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `let pvtTable = Worksheets.Item("Sheet1").Range("A3").PivotTable let strPgName = pvtTable.PivotFields("Country").CurrentPage.Name` |

#### **PivotField.CurrentPageList**

返回或设置对应于项目列表的字符串数组，该项目列表包含于数据透视表的多项目页字段中。**Variant** 类型，可读写。

**语法**

**express.CurrentPageList**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

为了避免运行时错误，数据源必须是 OLAP 源，选择的字段当前必须位于页面位置中，并且 **EnableMultiplePageItems** 属性必须设置为 **True**。

**示例**

本示例设置页字段以列出数据透视表的“Food”项目。本示例假定数据透视表位于活动工作表上。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `function UseCurrentPageList(){     let pvtTable = ActiveSheet.PivotTables(1)     let pvtField = pvtTable.PivotFields("[Product]") 									     // To avoid run-time errors set the following property to True.     pvtTable.CubeFields("[Product]").EnableMultiplePageItems = true 									     // Set the page list to "Food".     pvtField.CurrentPageList = "[Product].[All Products].[Food]" }` |

#### **PivotField.CurrentPageName**

返回或设置指定数据透视表上的当前显示页。该页名称将出现在页字段中。注意，只有当已存在当前显示页时，本属性才有效。**String** 类型，可读写。

**语法**

**express.CurrentPageName**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

本属性应用于与 OLAP 数据源相连的数据透视表。如果用未与 OLAP 数据源相连的数据透视表返回或设置本属性，则将导致运行时错误。

**示例**

本示例将活动工作表上第一个数据透视表的当前显示页的名称设置为“USA”。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `ActiveSheet.PivotTables("PivotTable1").PivotFields("[Customers]").CurrentPageName = "[Customers].[All Customers].[USA]"` |

#### **PivotField.DatabaseSort**

如果为 **True**，则允许手动更改数据透视表字段中项目的位置。如果该字段中没有手动定位的项目，则返回 **True**。**Boolean** 类型，可读写。

**语法**

**express.DatabaseSort**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

如果数据源不是联机分析处理 (OLAP)?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，则 **DatabaseSort** 属性返回 **False**。

如果数据源是 OLAP，并且字段中既没有应用自定义排序也没有应用自动排序，那么该属性返回 **True**。

对于 OLAP 数据透视表，如果将 **DatabaseSort** 属性设置为 **True**，则会删除应用于字段的所有自定义排序或自动排序（也就是说，建立连接时数据透视表恢复为默认的状态）。

如果没有应用自动排序，那么将 **DatabaseSort** 属性设置为 **False** 时，会使排序次序变为当前的项目次序。

将 **DatabaseSort** 属性设置为 **True** 或 **False** 都会引起更新。

对于非 OLAP 源或 OLAP 数据字段，如果将 **DatabaseSort** 属性设置为 **True**，则会导致运行时错误。

**示例**

本示例判断数据源是否是 OLAP 数据源，并通知用户。本示例假定 OLAP 数据透视表位于活动工作表上。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `function UseDatabaseSort(){     let pvtTable = ActiveSheet.PivotTables(1)     let pvtField = pvtTable.PivotFields("[Product].[Product Family]")                                          // Determine source type for the PivotTable report.     if(pvtField.DatabaseSort == true){         MsgBox("The source is OLAP; you can manually reorder items.")     }     else{         MsgBox("The data source might not be OLAP.")     } }` |

#### **PivotField.DataRange**

返回一个 **Range** 对象，如下表所示。只读。

**语法**

**express.DataRange**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

| 对象             | 数据区域               |
| ---------------- | ---------------------- |
| “数据”字段       | 字段中包含的数据       |
| 行、列或者页字段 | 字段中的数据项         |
| 数据项           | 数据项规范所限定的数据 |

**示例**

此示例选定名为“REGION”字段中的数据透视表数据项。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `let pvtTable = Worksheets.Item("Sheet1").Range("A3").PivotTable Worksheets.Item("Sheet1").Activate() pvtTable.PivotFields("REGION").DataRange.Select()` |

#### **PivotField.DataType**

返回一个 **XlPivotFieldDataType** 值，它代表数据透视表字段中的数据类型。

**语法**

**express.DataType**

*express*   一个代表 **PivotField** 对象的变量。

**示例**

本示例显示“ORDER_DATE”字段的数据类型。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `let pvtTable = Worksheets.Item("Sheet1").Range("A3").PivotTable switch(pvtTable.PivotFields("ORDER_DATE").**DataType**) {     case xlText:         MsgBox("The field contains text data")         break     case xlNumber:         MsgBox("The field contains numeric data")         break     case xlDate:         MsgBox("The field contains date data") }` |

#### **PivotField.DisplayAsCaption**

table { word-break:break-all; }

此属性用于将透视字段的成员属性显示为标题。只读。

**语法**

**express.DisplayAsCaption**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

table { word-break:break-all; }

当给定成员属性用作标题时，此属性返回 **True**，当透视字段的成员属性未用作标题时，返回 **False**。尝试将此属性用于并非成员属性的透视字段时，将返回运行时错误。

#### **PivotField.DisplayAsTooltip**

table { word-break:break-all; }

此属性用于指定在工具提示中是否显示透视字段的特定成员属性。可读写。**Boolean**。

**语法**

**express.DisplayAsTooltip**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

table { word-break:break-all; }

如果尝试从透视字段的非成员属性获取或设置这些属性，将返回运行时错误。

当 **DisplayAsTooltip** 属性设置为 **True** 时，将在工具提示中显示给定成员属性，当此属性设置为 **False** 时，不会在工具提示中显示给定成员属性。默认值为 **True**。

#### **PivotField.DisplayInReport**

table { word-break:break-all; }

此属性用于指定在数据透视表中是否显示指定的成员属性 PivotField。可读/写 **Boolean** 类型。

**语法**

**express.DisplayInReport**

*express*   一个代表 **PivotField** 对象的变量。

#### **PivotField.DragToColumn**

如果指定字段能被拖动到列位置上，则为 **True**。默认值是 **True**。**Boolean** 类型，可读写。

**语法**

**express.DragToColumn**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源的度量字段，该值为 **False**。

**示例**

**示例**

此示例防止将第一张工作表上第一个数据透视表中的“Year”字段拖动到列位置上。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `Worksheets.Item(1).PivotTables("Pivot1").PivotFields("Year").**DragToColumn** = false` |

#### **PivotField.DragToData**

如果指定字段可被拖动到数据位置上，则为 **True**。默认值是 **True**。**Boolean** 类型，可读写。

**语法**

**express.DragToData**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源的度量字段，该值为 **False**。

**示例**

此示例防止将“Year”字段拖动到第一张工作表上第一个数据透视表中的数据位置上。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `Worksheets.Item(1).PivotTables("Pivot1").PivotFields("Year").**DragToData** = false` |

#### **PivotField.DragToHide**

如果通过将字段拖离数据透视表可隐藏该字段，则为 **True**。默认值是 **True**。**Boolean** 类型，可读写。

**语法**

**express.DragToHide**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

**示例**

此示例防止将第一张工作表上第一个数据透视表中的“Year”字段拖离该报表。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `Worksheets.Item(1).PivotTables("Pivot1").PivotFields("Year").**DragToHide** = false` |

#### **PivotField.DragToPage**

如果字段可被拖动到页位置上，则为 **True**。默认值是 **True**。**Boolean** 类型，可读写。

**语法**

**express.DragToPage**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源的度量字段，该值为 **False**。

**示例**

此示例防止将第一张工作表上第一个数据透视表中的“Year”字段拖动到页位置上。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `Worksheets.Item(1).PivotTables("Pivot1").PivotFields("Year").**DragToPage** = false` |

#### **PivotField.DragToRow**

如果字段可被拖动到行位置上，则为 **True**。默认值是 **True**。**Boolean** 类型，可读写。

**语法**

**express.DragToRow**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源的度量字段，该值为 **False**。

**示例**

此示例防止第一张工作表上第一个数据透视表中的“Year”字段被拖动到行位置上。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `Worksheets.Item(1).PivotTables("Pivot1").PivotFields("Year").**DragToRow** = false` |

#### **PivotField.DrilledDown**

如果指定数据透视表字段或数据透视表项的标志设置为“drilled”（展开或可见），则为 **True**。**Boolean** 类型，可读写。

**语法**

**express.DrilledDown**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

此属性仅对 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源有效。

如果字段或项被隐藏，则不可设置此属性。

**示例**

此示例将活动工作表上第三个数据透视表中状态字段的所有项的标记设置为“not drilled”。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `ActiveSheet.PivotTables("PivotTable3").PivotFields("state").**DrilledDown** = false` |

#### **PivotField.EnableItemSelection**

如果为 **False**，则在用户界面中禁止使用下拉字段的功能。默认值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.EnableItemSelection**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

如果 OLAP 数据透视表字段不是层次结构中的最高级别，则将出现运行时错误。

**示例**

**示例**

本示例确定使用下拉字段的选定项目的设置，并且在必要时启用该功能。本示例假定数据透视表位于活动工作表上。
示例代码
function UseEnableItemSelection() {
let pvtTable = ActiveSheet.PivotTables(1)
let pvtField = pvtTable.RowFields(1)
// Determine setting for property and enable if necessary.
if(pvtField.EnableItemSelection == false) {
pvtField.EnableItemSelection = true
MsgBox("Item selection enabled for fields.")
}
else {
MsgBox("Item selection is already enabled for fields.")
}
}

#### **PivotField.EnableMultiplePageItems**

table { word-break:break-all; }

用于指定对于页面区域中的字段是否在筛选器下拉列表中显示复选框。可读/写 **Boolean** 类型。

**语法**

**express.EnableMultiplePageItems**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

table { word-break:break-all; }

为 OLAP 保留现有属性值。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在 ET 2007 或更高版本中，如果创建先于 ET 2007 的 OLAP 数据透视表 (PivotTable.Version < 3)，且 **PivotTable** 对象的 **SubtotalHiddenPageItems** 属性和 **PivotField** 对象的 **EnableMultiplePageItems** 属性均设置为 **True**，则更改页面区域的筛选器下拉菜单中的复选框的状态将不会有效果。在这种情况下，筛选器将始终设置为 **All**，并包含未选中（隐藏）的项目。 |

#### **PivotField.Formula**

table { word-break:break-all; }

返回或设置一个 **String** 值，它代表 A1 样式表示法和宏语言中的对象的公式。

**语法**

**express.Formula**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

table { word-break:break-all; }

此属性对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源无效。

如果单元格包含一个常量，此属性返回该常量。如果单元格为空，此属性返回空字符串。如果单元格包含公式，**Formula** 属性将该公式作为字符串返回，所用格式与在编辑栏（包括等号）中显示时的格式相同。

如果将单元格的值或者公式设置为日期类型，则 ET 将检查此单元格的数字格式是否符合日期或者时间格式。如果不符合，ET 将把数字格式设置为默认的短日期格式。

如果指定区域是一维或二维区域，则可将公式指定为 Visual Basic 中相同维数的数组。同样，也可在 Visual Basic 数组中使用公式。

如果为多单元格区域设置公式，则会用公式填充该区域所有的单元格。

#### **PivotField.Function**

返回或设置对数据透视表字段汇总时所使用的函数（仅用于数据字段）。**XlConsolidationFunction** 类型，可读写。

**语法**

**express.Function**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，本属性是只读的，并且总是返回 **xlUnknown**。对于其他的数据源，本属性不能被设置为 **xlUnknown**。

**示例**

本示例对活动工作表中第一个数据透视表的“Sum of 1994”字段进行设置，使之使用“SUM”函数。

#### **PivotField.GroupLevel**

返回一组字段中指定字段的位置（如果该字段是分组字段集合中的成员）。只读。

**语法**

**express.GroupLevel**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

该属性对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源无效。

处于最高级的父字段（即最左侧的父字段）为第一级，其子字段为第二级，依此类推。

**示例**

本示例检查包含活动单元格的字段是否为最高一级的父字段，如果是则显示一个消息框。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `Worksheets.Item("Sheet1").Activate() if(ActiveCell.PivotField.**GroupLevel** == 1) {     MsgBox("This is the highest-level parent field.") }` |

#### **PivotField.Hidden**

table { word-break:break-all; }

此属性用于隐藏 OLAP 层次结构的各个级别。可读/写 **Boolean** 类型。

**语法**

**express.Hidden**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

table { word-break:break-all; }

对于隐藏的级别，返回 **True**；对于可见（非隐藏）的级别，返回 **False**。

此属性只适用于 OLAP 层次结构的级别。

对于关系数据源和 OLAP 特性，此属性始终为 **False**。如果试图为特性或关系字段将其设置为 **True**，将会返回运行时错误。 如果此设置对于同一层次结构的所有其他级别已设置为 **True**，则无法为某个级别将此属性设置为 **True**。试图这样做将会返回运行时错误。必须为层次结构中的至少一个级别将此属性设置为 **False**。

#### **PivotField.HiddenItems**

返回一个对象，该对象表示单个隐藏数据透视表项（**PivotItem** 对象），或表示指定字段中所有隐藏项的集合（**PivotItems** 对象）。只读。

**语法**

**express.HiddenItems**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，该属性总返回一个空集合。

**示例**

本示例将“product”字段的所有隐藏数据项的名称的列表添加到一张新工作表中。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `let nwSheet = Worksheets.Add() nwSheet.Activate() let pvtTable = Worksheets.Item("Sheet2").Range("A1").PivotTable let rw = 0 for(let i=1;i <= pvtTable.PivotFields("product").**HiddenItems**.Count;i++) {     rw++     nwSheet.Cells.Item(rw, 1).Value2 = pvtTable.PivotFields("product").HiddenItems.Item(i).Name }` |

#### **PivotField.HiddenItemsList**

返回或设置一个字符串数组，该数组为数据透视表字段的隐藏项。**Variant** 类型，可读写

**语法**

**express.HiddenItemsList**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

**HiddenItemsList** 属性只对联机分析处理 (OLAP)?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源有效，如果将该属性用于非 OLAP 数据源，则将返回一个运行错误。

**示例**

本示例设置项目列表，从而只显示某些项目。本示例假定 OLAP 数据透视表位于活动工作表上。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `function UseHiddenItemsList() {     ActiveSheet.PivotTables(1).PivotFields(1).**HiddenItemsList** = ["[Product].[All Products].[Food]","[Product].[All Products].[Drink]"] }` |

#### **PivotField.IncludeNewItemsInFilter**

table { word-break:break-all; }

在将手动筛选应用于透视字段时，此属性允许开发人员指定是应跟踪排除的项目还是应跟踪包含的项目。**Boolean** 类型，可读写。

**语法**

**express.IncludeNewItemsInFilter**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

table { word-break:break-all; }

此属性的默认值为 **False**。

应用手动筛选后，开发人员可以将 **IncludeNewItemsInFilter** 属性设置为 **True** 来跟踪排除的项目，也可以将此属性设置为 **False** 来跟踪包含的项目。如果在 **IncludeNewItemsInFilter** 属性设置为 **True** 后向源数据中添加了新项目，则这些新项目只有在进行新的刷新操作后才会显示在数据透视表中，因为它们未处在 ET 正在跟踪的项目集合中。

对于 OLAP 层次结构，此设置在 **CubeField** 对象上设置，而试图在作为层次结构一部分的透视字段上设置它将会失败（并会产生运行时错误）。可以通过作为层次结构一部分的透视字段获取此设置，而且它将返回与对应的 **CubeField.IncludeNewItemsInFilter** 属性相同的结果。

切换此设置后，以下集合将为空：**HiddenItemsList**、**HiddenItems**、**VisibleItemsList** 和 **VisibleItems**。 当 **IncludeNewItemsInFilter** 设置为 **False** 时，**HiddenItemsList** 和 **HiddenItems** 集合为空，并且无法向它们添加项目。试图添加项目时会返回运行时错误。 当 **IncludeNewItemsInFilter** 设置为 **True** 时，**VisibleItemsList** 和 **VisibleItems** 集合为空，并且无法向它们添加项目。试图添加项目时会返回运行时错误。

在数据透视表中，此设置在 **PivotField** 对象上设置。切换此设置将不会改变筛选器状态。

#### **PivotField.IsCalculated**

如果数据透视表字段为计算字段或计算项，则此属性为 **True**。**Boolean** 类型，只读。

**语法**

**express.IsCalculated**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，此属性始终返回 **False**。

**示例**

如果指定数据透视表包含任何计算字段，此示例将禁用**“数据透视表字段”**对话框。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `let pt = Worksheets.Item(1).PivotTables("Pivot1") for(let i=1;i <= pt.PivotFields().Count;i++) {     if(pt.PivotFields(i).**IsCalculated**) {         pt.EnableFieldDialog = false     } }` |

#### **PivotField.IsMemberProperty**

如果数据透视表字段包含成员属性，则返回 **True**。**Boolean** 类型，只读。

**语法**

**express.IsMemberProperty**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

如果没有使用联机分析处理 (OLAP)?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，则该属性将返回一个运行时错误。

**示例**

本示例确定数据透视表字段是否包含成员属性，并通知用户。本示例假定数据透视表位于活动的工作表上，并与 OLAP 数据源相连。
示例代码
function CheckForMembers() {
let pvtTable = ActiveSheet.PivotTables(1)
let pvtField = pvtTable.PivotFields(1)
// Determine if member properties exist and notify user.
if(pvtField.IsMemberProperty == true) {
MsgBox("The PivotField contains member properties.")
}
else {
MsgBox("The PivotField does not contain member properties.")
}
}

#### **PivotField.LabelRange**

返回一个 **Range** 对象，它代表包含字段标签的单元格。只读。

**语法**

**express.LabelRange**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

**示例**

此示例选择“ORDER_DATE”字段的字段按钮。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `let pvtTable = Worksheets.Item("Sheet1").Range("A3").PivotTable let pvtField = pvtTable.PivotFields("ORDER_DATE") Worksheets.Item("Sheet1").Activate() pvtField.**LabelRange**.Select()` |

#### **PivotField.LayoutBlankLine**

如果在数据透视表的指定行字段后插入了一个空行，则该值为 **True**。默认值为 **False**。**Boolean** 类型，可读写。

**语法**

**express.LayoutBlankLine**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

可以对任意数据透视表字段设置此属性；然而，只有当指定字段是除最里层（最低级别?（级别：OLAP 维的一部分。在维中，数据的明细级别会有高、低之分，如在时间维内分为年、季度、月和日等级别。））行字段之外的其他行字段时，该空行才会出现。对于非 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，当重新排列该字段或将其添加到数据透视表中时，此属性的值不会改变。

不能在数据透视表的空行中输入数据。

**示例**

本示例在活动工作表的第一个数据透视表的状态字段后添加一个空行。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `ActiveSheet.PivotTables("PivotTable1").PivotFields("state").**LayoutBlankLine** = true` |

#### **PivotField.LayoutCompactRow**

table { word-break:break-all; }

指定在选择行时是否压缩透视字段（在一列中显示多个透视字段的项目）。可读/写 **Boolean** 类型。

**语法**

**express.LayoutCompactRow**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

table { word-break:break-all; }

如果 **LayoutForm** 属性设置为 **xlTabular**，则将此属性设置为 **True**（默认值）将不会立即在数据透视表中产生可看到的效果，但可以独立地设置此属性。

#### **PivotField.LayoutForm**

返回或设置指定的数据透视表项出现的方式，即以表格格式还是以分级显示格式显示。**XlLayoutFormType** 类型，可读写。

**语法**

**express.LayoutForm**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

| **XlLayoutFormType** 可为以下 **XlLayoutFormType** 常量之一。 |
| ------------------------------------------------------------ |
| **xlTabular**。默认值。                                      |
| **xlOutline**。**LayoutSubtotalLocation** 属性指定分类汇总在数据透视表中出现的位置。 |

可以对任意数据透视表字段设置此属性。然而，只有当指定字段是除最里层（最低级别?（级别：OLAP 维的一部分。在维中，数据的明细级别会有高、低之分，如在时间维内分为年、季度、月和日等级别。））行字段之外的其他行字段时，该格式才会出现。对于非 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，当重新排列该字段或将其添加到数据透视表中或从数据透视表中删除时，此属性的值不会更改。

**示例**

本示例以分级显示格式显示当前活动工作表上第一个数据透视表中的状态字段，然后在该字段顶部显示分类汇总。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `ActiveSheet.PivotTables("PivotTable1").PivotFields("state").**LayoutForm** = xlOutline ActiveSheet.PivotTables("PivotTable1").PivotFields("state").LayoutSubtotalLocation = xlTop` |

#### **PivotField.LayoutPageBreak**

**语法**

**express.LayoutPageBreak**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

虽然可以对任意数据透视表字段设置此属性，但只有当指定字段是除最里层（最低级别?（级别：OLAP 维的一部分。在维中，数据的明细级别会有高、低之分，如在时间维内分为年、季度、月和日等级别。））行字段之外的其他行字段时，打印选项才会出现。对于非 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，当重新排列该字段或将其添加到数据透视表中或从数据透视表中删除时，此属性的值不会更改。

**示例**

本示例在活动工作表的第一个数据透视表的状态字段之后添加一个分页符。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `ActiveSheet.PivotTables("PivotTable1").PivotFields("state").**LayoutPageBreak** = true` |

#### **PivotField.LayoutSubtotalLocation**

返回或设置与指定字段相关（在其上面或下面）的数据透视表字段分类汇总的位置。**XlSubtototalLocationType** 类型，可读写。

**语法**

**express.LayoutSubtotalLocation**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

可以对任意具有分级显示格式的数据透视表字段设置此属性；但是，只有当指定字段是除最里层（最低级别?（级别：OLAP 维的一部分。在维中，数据的明细级别会有高、低之分，如在时间维内分为年、季度、月和日等级别。））行字段之外的其他行字段时，格式才会出现。对于非 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，当重新排列该字段或将其添加到数据透视表中或从数据透视表中删除时，此属性的值不会更改。

**LayoutForm** 属性确定报表是以表格格式还是以分级显示格式显示。

**示例**

本示例以分级显示格式显示当前活动工作表上第一个数据透视表中的状态字段，然后在该字段顶部显示分类汇总。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `ActiveSheet.PivotTables("PivotTable1").PivotFields("state").LayoutForm = xlOutline ActiveSheet.PivotTables("PivotTable1").PivotFields("state").**LayoutSubtotalLocation** = xlAtTop` |

#### **PivotField.MemberPropertyCaption**

table { word-break:break-all; }

设置 **MemberPropertyCaption** 属性可控制将哪个成员属性用作指定级别的标题。可读写 **Boolean** 类型。

**语法**

**express.MemberPropertyCaption**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

table { word-break:break-all; }

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 只有在为透视字段将 **UseMemberPropertyAsCaption** 设置为 **True** 时，此设置才会产生可看到的效果。 |

设置 **MemberPropertyCaption** 后，在切换 **UseMemberPropertyAsCaption** 属性时会记住此设置。

#### **PivotField.MemoryUsed**

返回对象当前使用的内存大小，以字节表示。**Long** 型，只读。

**语法**

**express.MemoryUsed**

*express*   一个代表 **PivotField** 对象的变量。

**示例**

此示例弹出一个消息框，该消息框中显示了 ET 当前使用的内存字节数。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `MsgBox("ET is currently using " + Application.**MemoryUsed** + " bytes")` |

#### **PivotField.Name**

table { word-break:break-all; }

返回或设置一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

*express*   一个代表 **PivotField** 对象的变量。

#### **PivotField.NumberFormat**

table { word-break:break-all; }

返回或设置一个 **String** 值，它代表对象的格式代码。

**语法**

**express.NumberFormat**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

table { word-break:break-all; }

只能对数据字段设置 **NumberFormat** 属性。

格式代码与**“设置单元格格式”**对话框中的**“格式代码”**选项是同一个字符串。**Format** 函数使用的格式代码字符串与 **NumberFormat** 和 **NumberFormatLocal** 属性使用的格式代码字符串不同。

#### **PivotField.Orientation**

返回或设置一个 **XlPivotFieldOrientation** 值，它代表字段在指定的数据透视表中的位置。

**语法**

**express.Orientation**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，当设置某一层中的一个字段的此属性值时，也会同时设置同一层中所有其他字段的方向。维字段只能在数据透视表的行、列和页字段区域中进行定向。度量字段则只能在数据区域中进行定向。将某一层或数据字段设置为 **xlHidden** 时，会将该层或字段从数据透视表中移出。

**示例**

本示例显示“ORDER_DATE”字段的方向。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `let pvtTable = Worksheets.Item("Sheet1").Range("A3").PivotTable let pvtField = pvtTable.PivotFields("ORDER_DATE") switch(pvtField.**Orientation**) {     case xlHidden:         MsgBox("Hidden field")         break     case xlRowField:         MsgBox("Row field")         break     case xlColumnField:         MsgBox("Column field")         break     case xlPageField:         MsgBox("Page field")         break     case xlDataField:         MsgBox("Data field") }` |

#### **PivotField.Parent**

table { word-break:break-all; }

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **PivotField** 对象的变量。

#### **PivotField.ParentField**

返回一个 **PivotField** 对象，该对象表示作为指定对象的组父级的数据透视表字段。只读。

**语法**

**express.ParentField**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

**示例**

本示例显示包含活动单元格的字段的组父级字段名称。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `Worksheets.Item("Sheet1").Activate() MsgBox("The active field is a child of the field " + ActiveCell.PivotField.**ParentField**.Name)` |

#### **PivotField.ParentItems**

返回一个对象，该对象表示指定字段中的单个数据透视表项（**PivotItem** 对象），或者作为组父级的所有项的一个集合（**PivotItems** 对象）。指定字段必须是另一个字段的组父级。只读。

**语法**

**express.ParentItems**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

该属性不可用于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源。

**示例**

本示例创建“product”字段中所有组父项的名称的列表。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `let nwSheet = Worksheets.Add() nwSheet.Activate() let pvtTable = Worksheets.Item("Sheet2").Range("A1").PivotTable let rw = 0 for(let i=1;i <= pvtTable.PivotFields("product").**ParentItems**.Count;i++) {     rw++     nwSheet.Cells.Item(rw, 1).Value2 = pvtTable.PivotFields("product").ParentItems.Item(i).Name }` |

#### **PivotField.PivotFilters**

table { word-break:break-all; }

返回或设置指定 **PivotField** 对象的 PivotFilters。只读。

**语法**

**express.PivotFilters**

*express*   一个代表 **PivotField** 对象的变量。

#### **PivotField.Position**

table { word-break:break-all; }

返回一个 Variant 值，他代表数据透视表上所有字段的位置，包括属于所有行、列、页以及数据方向上的字段（第一字段、第二字段、第三字段等等）。

**语法**

**express.Position**

*express*   一个代表 **PivotField** 对象的变量。

#### **PivotField.PropertyOrder**

只对属于成员属性字段的数据透视表字段有效。返回一个 **Long** 类型的数值，该数值表示成员属性在其所属的多维数据集字段内的显示位置。可读写。

**语法**

**express.PropertyOrder**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

设置该属性将重新安排该多维数据集字段的属性的顺序。该属性从一开始。允许的区域从一到为层次结构显示的成员属性字段的最大数量。

如果 **IsMemberProperty** 属性为 **False**，则使用 **PropertyOrder** 属性将导致一个运行时错误。

**示例**

本示例判断第四个字段中是否存在成员属性，如果存在，则显示成员属性的位置。 ET 将该信息通知用户。本示例假定数据透视表位于活动工作表上，并且基于联机分析处理 (OLAP)?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源。
示例代码
function CheckPropertyOrder() {
let pvtTable = ActiveSheet.PivotTables(1)
let pvtField = pvtTable.PivotFields(4)
// Check for member properties and notify user.
if(pvtField.IsMemberProperty == false) {
MsgBox("No member properties present.")
}
else {
MsgBox("The property order of the members is: " + pvtField.PropertyOrder)
}
}

#### **PivotField.PropertyParentField**

返回一个 **PivotField** 对象，该对象表示字段中属性所从属的字段。

**语法**

**express.PropertyParentField**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

只对属于成员属性字段的字段有效。

如果 **IsMemberProperty** 属性为 **False**，则使用 **PropertyParentField** 属性将返回一个运行时错误。

**示例**

本示例确定第四个字段中是否存在成员属性，如果存在，则确定属性所属的字段。 ET 将该信息通知用户。本示例假定数据透视表位于活动工作表上，并且基于联机分析处理 (OLAP)?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源。
示例代码
function CheckParentField() {
let pvtTable = ActiveSheet.PivotTables(1)
let pvtField = pvtTable.PivotFields(4)
// Check for member properties and notify user.
if(pvtField.IsMemberProperty == false) {
MsgBox("No member properties present.")
}
else {
MsgBox("The parent field of the members is: " + pvtField.PropertyParentField.Name)
}
}

#### **PivotField.RepeatLabels**

table { word-break:break-all; }

返回或设置在数据透视表中是否对指定的透视字段重复项目标签。可读/写

**语法**

**express.RepeatLabels**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

table { word-break:break-all; }

如果对指定的透视字段重复项目标签，则为 **True**；否则为 **False**。

**RepeatLabels** 属性的设置对应于数据透视表中的字段的**“字段设置”**对话框**“布局和打印”**选项卡上的**“重复项目标签”**复选框。

若要指定是否在单次操作中对数据透视表内的所有透视字段重复项目标签，请使用 **RepeatAllLabels** 方法。

#### **PivotField.ServerBased**

如果指定数据透视表的数据源为外部数据源，并且只检索与选定页字段相匹配的数据项，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ServerBased**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

该属性不能应用于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，这时其值总是 **False**。

如果本属性设为 **True**，则指定数据库中只有与选定页字段项相匹配的记录可被检索到。以后每当用户更改选定页字段时，新的选定页字段项将作为参数传递给查询，并且高速缓存将得到刷新。

如果下列某个条件成立，则不能对该属性进行设置：

- 字段被分组。
- 使用的不是外部数据源。
- 高速缓存由两个或者多个数据透视表共享使用。
- 字段的数据类型是不能基于服务器的类型（即备注字段或 OLE 对象）。

**示例**

本示例列出所有基于服务器的页字段。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `let r = 0 for(let i=1;i <= ActiveSheet.PivotTables(1).PageFields().Count;i++) {     if(ActiveSheet.PivotTables(1).PageFields(i).**ServerBased** == true) {         r++         Worksheets.Item(2).Cells.Item(r, 1).Value2 = ActiveSheet.PivotTables(1).PageFields(i).Name     } }` |

#### **PivotField.ShowAllItems**

如果显示数据透视表中的所有项目（即使这些项目中不包含汇总数据），则该值为 **True**。默认值为 **False**。**Boolean** 类型，可读写。

**语法**

**express.ShowAllItems**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，该值总是 **False**。

**示例**

本示例显示第一张工作表上第一张数据透视表中“Month”字段的所有行，包括没有数据的月份。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `Worksheets.Item(1).PivotTables("Pivot1").PivotFields("Month").**ShowAllItems** = true` |

#### **PivotField.ShowDetail**

table { word-break:break-all; }

获取或设置指定的 **PivotField** 是否显示明细数据。可读/写 **Boolean** 类型。

**语法**

**express.ShowDetail**

*express*   一个代表 **PivotField** 对象的变量。

#### **PivotField.ShowingInAxis**

table { word-break:break-all; }

指明透视字段当前在数据透视表中是否可见。只读。

**语法**

**express.ShowingInAxis**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

table { word-break:break-all; }

如果透视字段当前可见，则此属性返回 **True**。如果透视字段当前不可见，则返回 **False**。透视字段的可见性取决于透视字段标题是否在表格（非压缩）形式的数据透视表中可见。

#### **PivotField.SourceCaption**

table { word-break:break-all; }

**SourceCaption** 属性只适用于 OLAP 数据透视表，并从 OLAP 服务器中返回透视字段的原始标题。只读。

**语法**

**express.SourceCaption**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

table { word-break:break-all; }

可以将 **PivotField**.**Name** 或 **PivotField**.**SourceName** 用于非 OLAP 数据透视表。

#### **PivotField.SourceName**

table { word-break:break-all; }

返回一个 **String** 值，它代表指定的对象出现在指定的数据透视表的原始源数据中时的名称。

**语法**

**express.SourceName**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

table { word-break:break-all; }

如果用户在创建数据透视表之后重命名数据项，此属性的值可能会与当前数据项名称不同。

下表显示 **SourceName** 属性及其相关属性的示例值，假设存在唯一名称为“[Europe].[France].[Paris]”的 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，并且存在包含名为“Paris”的项的非 OLAP 数据源。

| 属性           | 值（OLAP 数据源）                 | 值（非 OLAP 数据源）        |
| -------------- | --------------------------------- | --------------------------- |
| **Caption**    | Paris                             | Paris                       |
| **Name**       | [Europe].[France].[Paris]（只读） | Paris                       |
| **SourceName** | [Europe].[France].[Paris]（只读） | （与 SQL 属性值相同，只读） |
| **Value**      | [Europe].[France].[Paris]（只读） | Paris                       |

在将一个索引值指定到 **PivotItems** 集合中时，您可以使用下表所示的语法。

| 语法（OLAP 数据源）                                | 语法（非 OLAP 数据源）         |
| -------------------------------------------------- | ------------------------------ |
| expression.PivotItems("[Europe].[France].[Paris]") | expression.PivotItems("Paris") |

在使用 **Item** 属性引用集合中的特定成员时，可以使用下表所示的文本索引名称。

| 名称（OLAP 数据源）       | 名称（非 OLAP 数据源） |
| ------------------------- | ---------------------- |
| [Europe].[France].[Paris] | Paris                  |

#### **PivotField.StandardFormula**

返回或设置一个 **String** 值，该值指定使用标准英语（美国）格式的公式。可读写。

**语法**

**express.StandardFormula**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

**StandardFormula** 属性主要影响具有日期或数字格式的项目名称。该属性提供了一种方法可指定或查询给定计算项的公式。

**StandardFormula** 属性是国际通用的，**Formula** 属性则不是。

**示例**

本示例向“Decimals”字段中添加 10，并将其显示为数据字段中的计算项。本示例假定数据透视表位于活动工作簿上，并且标题为“Decimals”的字段位于模拟运算表中。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `function UseStandardFomula() {     let pvtTable = ActiveSheet.PivotTables(1)     // Change calculated field of decimals by adding '10'.     pvtTable.CalculatedFields().Item(1).**StandardFormula** = "Decimals + 10" }` |

#### **PivotField.SubtotalName**

返回或设置显示在指定数据透视表的分类汇总列或行标题中的文本字符串标志。默认值为“Subtotal”。**String** 类型，可读写。

**语法**

**express.SubtotalName**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

**示例**

本示例将活动工作表上第二个数据透视表的状态字段中的分类汇总标志设置为“Regional Subtotal”（取代默认的“Subtotal”）。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `ActiveSheet.PivotTables("PivotTable2").PivotFields("state").**SubtotalName** = "Regional Subtotal"` |

#### **PivotField.Subtotals**

返回或设置与指定字段同时显示的分类汇总。仅对非数据字段有效。**Variant** 类型，可读写。

**语法**

**express.Subtotals**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

如果索引为 **True**，则该字段显示分类汇总。如果索引 1（自动）为 **True**，则其他所有值将设置为 **False**。

| 索引 | 含义       |
| ---- | ---------- |
| 1    | 自动       |
| 2    | 总计       |
| 3    | 计数       |
| 4    | 平均       |
| 5    | 最大值     |
| 6    | 最小值     |
| 7    | 乘积       |
| 8    | Count Nums |
| 9    | StdDev     |
| 10   | StdDevp    |
| 11   | Var        |
| 12   | Varp       |

对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，*Index* 只能返回或设置为 1（自动）。返回的数组对第一个数组元素总是包含 **True** 或 **False**，而对所有其他元素则包含 **False**。数组中所有元素的值都为 **False**，则表明没有分类汇总。

**示例**

本示例将包含活动单元格的字段设置为显示分类汇总求和。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `Worksheets.Item("Sheet1").Activate() ActiveCell.PivotField.**Subtotals(2, true)**` |

#### **PivotField.TotalLevels**

返回当前字段组中的字段总数。如果字段没有分组或数据源基于的是 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。），则 **TotalLevels** 将返回值 1。**Long** 类型，只读。

**语法**

**express.TotalLevels**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

同一分组内的所有字段具有相同的 **TotalLevels** 值。

**示例**

本示例显示包含活动单元格的字段组中字段的总数。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `Worksheets.Item("Sheet1").Activate() MsgBox("This group has " + ActiveCell.PivotField.**TotalLevels** + " levels")` |

#### **PivotField.UseMemberPropertyAsCaption**

table { word-break:break-all; }

此属性用于控制是否将成员属性标题用于透视字段的 PivotItem 标题。可读/写 **Boolean** 类型。

**语法**

**express.UseMemberPropertyAsCaption**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

table { word-break:break-all; }

如果将透视字段的 **UseMemberPropertyAsCaption** 设置为 **True**，则 **MemberPropertyCaption** 将指定要显示的成员属性标题。如果未指定，则该透视字段的第一个成员属性（按数据源顺序）将显示为该透视字段的项目的标题。

如果将 **UseMemberPropertyAsCaption** 设置为 **False**，则常规 PivotItem 标题将用于该透视字段。

如果尝试将没有成员属性的透视字段的 **UseMemberPropertyAsCaption** 设置为 **True**，则会返回一个运行时错误。对于没有成员属性的透视字段，该属性将始终为 **False**。

#### **PivotField.Value**

table { word-break:break-all; }

返回或设置一个 **String** 值，它代表数据透视表中指定的字段的名称。

**语法**

**express.Value**

*express*   一个代表 **PivotField** 对象的变量。

#### **PivotField.VisibleItems**

返回一个对象，该对象表示指定字段中的单个可见数据透视表项（**PivotItem** 对象），或指定字段中的所有可见项的集合（**PivotItems** 对象）。只读。

**语法**

**express.VisibleItems**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，本属性是只读的，并且总返回 **True**。没有隐藏项。

**示例**

本示例向新工作表上的列表中添加“Product”字段中所有可见的数据项名称。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `let nwSheet = Worksheets.Add() nwSheet.Activate() let pvtTable = Worksheets.Item("Sheet2").Range("A1").PivotTable let rw = 0 for(let i=1;i <= pvtTable.PivotFields("Product").**VisibleItems**.Count;i++) {     rw++     nwSheet.Cells.Item(rw, 1).Value2 = pvtTable.PivotFields("Product").VisibleItems.Item(i).Name }` |

#### **PivotField.VisibleItemsList**

返回或设置一个 **Variant** 类型的值，该值指定一个字符串数组，字符串代表应用于透视字段的手动筛选中的包含项。可读/写。

**语法**

**express.VisibleItemsList**

*express*   一个代表 **PivotField** 对象的变量。

**说明**

此属性只适用于 OLAP 数据透视表。

**示例**

此示例显示 OLAP 数据透视表中包含在内的手动筛选。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `ActiveSheet.PivotTables("PivotTable2").PivotFields("[Customer].[Customer Geography].[Country]").VisibleItemsList = ["[Customer].[Customer Geography].[Country].[Australia]"] ActiveSheet.PivotTables("PivotTable2").PivotFields("[Customer].[Customer Geography].[State-Province]").VisibleItemsList = [""] ActiveSheet.PivotTables("PivotTable2").PivotFields("[Customer].[Customer Geography].[City]").VisibleItemsList = [""] ActiveSheet.PivotTables("PivotTable2").PivotFields("[Customer].[Customer Geography].[Postal Code]").VisibleItemsList = [""] ActiveSheet.PivotTables("PivotTable2").PivotFields("[Customer].[Customer Geography].[Full Name]").VisibleItemsList = [""]` |

适用环境：web

适用平台：windows/linux