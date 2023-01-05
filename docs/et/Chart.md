**Chart**



代表工作簿中的图表。

**说明**

此图表既可以是嵌入的图表（包含在 **ChartObject** 对象中），也可以是单独的图表工作表。

示例部分中描述了以下用于返回 **Chart** 对象的属性和方法：

- **Charts** 方法
- **ActiveChart** 属性
- **ActiveSheet** 属性

## 示例

**Charts** 集合包含工作簿中每个图表工作表的 **Chart** 对象。使用 **Charts**( *index*) 可以返回单个 **Chart** 对象，其中 index 为图表工作表的索引号或名称。图表索引号表示图表工作表在工作簿标签栏上的位置。*Charts(1)* 是工作簿中第一个（最左边的）图表；*Charts(Charts.Count)* 是最后一个（最右边的）图表。所有图表工作表均包括在索引计数中，即便是隐藏工作表也是如此。图表工作表名称显示在图表工作簿标签上。您可以使用 **Name** 属性设置或返回图表名称。本示例将更改第一张图表工作表中第一个系列的颜色。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Charts.Item(1).SeriesCollection(1).Format.Fill.ForeColor.RGB = rgbRed` |

本示例将名为“Sales”的图表移至活动工作簿的尾部。

| 示例代码复制                                                |
| ----------------------------------------------------------- |
| `Charts.Item("Sales").Move(null,Sheets.Item(Sheets.Count))` |

**Chart** 对象也是 **Sheets** 集合的成员，此集合包含工作簿中的所有工作表（图表工作表和工作表）。使用 **Sheets**(*index*) 可以返回单个工作表，其中 *index* 是工作表索引号或名称。

当图表是活动对象时，您可以使用 **ActiveChart** 属性引用它。如果用户选择了图表工作表，或者用 **Chart** 对象的 **Activate** 方法或 **ChartObject** 对象的 **Activate** 方法激活了它，则该图表工作表处于活动状态。本示例将激活图表工作表 1，然后设置图表类型和名称。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Charts.Item(1).Activate() let chart = ActiveChart     chart.Type = xlLine     chart.HasTitle = true     chart.ChartTitle.Text = "January Sales"` |

如果用户选择了嵌入图表，或者用 **Activate** 方法激活了包含该嵌入图表的 **ChartObject** 对象，则该嵌入图表处于活动状态。本示例将激活工作表 1 上的嵌入图表 1，然后设置图表类型和名称。请注意，在激活了嵌入图表之后，本示例中的代码与前一个示例中的代码相同。通过使用 **ActiveChart** 属性，您可以编写能够引用嵌入图表或图表工作表（视哪一个处于活动状态而定）的 Visual Basic 代码。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Worksheets.Item(1).ChartObjects(1).Activate() ActiveChart.ChartType = xlLine ActiveChart.HasTitle = true ActiveChart.ChartTitle.Text = "January Sales"` |

当图表工作表为活动工作表时，可以使用 **ActiveSheet** 属性来引用它。本示例使用 **Activate** 方法激活名为 Chart1 的图表工作表，然后将图表中系列 1 的内部颜色设置为蓝色。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Charts.Item("chart1").Activate() ActiveSheet.SeriesCollection(1).Format.Fill.ForeColor.RGB = rgbBlue` |

**方法**

|                                                              | 名称                     | 说明                                                         |
| ------------------------------------------------------------ | ------------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Activate**             | 使当前图表成为活动图表。                                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ApplyChartTemplate**   | 将标准图表类型或自定义图表类型应用于图表。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ApplyDataLabels**      | 将数据标签应用到图表中的所有系列。                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ApplyLayout**          | 应用功能区中显示的版式。                                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Axes**                 | 返回一个代表图表上单个坐标轴或坐标轴集合的某个对象。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ChartGroups**          | 返回一个对象，该对象表示图表中单个图表组（**ChartGroup** 对象）或所有图表组的集合（**ChartGroups** 对象）。返回的集合中包括每种类型的图表组。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ChartObjects**         | 返回一个对象，它代表工作表上的一个嵌入式图表（**ChartObject** 对象）或所有嵌入式图表的集合（**ChartObjects** 对象）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ChartWizard**          | 修改给定图表的属性。可使用本方法快速设置图表的格式，而不必逐个设置所有属性。本方法是非交互式的，并且仅更改指定的属性。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CheckSpelling**        | 检查对象的拼写。                                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ClearToMatchStyle**    | 清除图表元素格式以改为自动格式。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Copy**                 | 将工作表复制到工作簿的另一位置。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CopyPicture**          | 将所选对象作为图片复制到剪贴板。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**               | 删除对象。                                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Evaluate**             | 将一个 ET 名称转换为一个对象或者一个值。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Export**               | 以图形格式导出图表。                                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ExportAsFixedFormat**  | 导出为指定格式的文件。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **GetChartElement**      | 返回指定的 X 坐标和 Y 坐标上图表元素的信息。本方法稍有与众不同之处：调用时只须指定前两个参数，在本方法执行期间，ET 为其余参数赋值，本方法返回后应检验这些参数的值。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Location**             | 将图表移动到新位置。                                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Move**                 | 将图表移到工作簿的另一位置。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **OLEObjects**           | 返回一个对象，它代表图表或工作表上的单个 OLE 对象 (**OLEObject** ) 或所有 OLE 对象的集合（**OLEObjects** 集合）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Paste**                | 将剪贴板中的图表数据粘贴到指定的图表中。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PrintOut**             | 打印对象。                                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PrintPreview**         | 按对象打印后的外观效果显示对象的预览。                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Protect**              | 保护图表使其不被修改。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Refresh**              | 立即重新绘制指定的图表。                                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SaveAs**               | 将对图表或工作表的更改保存到另一不同文件中。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SaveChartTemplate**    | 向可用图表模板的列表添加自定义图表模板。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Select**               | 选择对象。                                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SeriesCollection**     | 返回一个对象，它代表图表或图表组中的一个系列（**Series** 对象）或所有系列的集合（**SeriesCollection** 集合）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetBackgroundPicture** | 为图表设置背景图形。                                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetDefaultChart**      | 指定 ET 新建图表时使用的图表模板的名称。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetElement**           | 设置图表上的图表元素。可读/写 **MsoChartElementType** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetSourceData**        | 为指定图表设置源数据区域。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Unprotect**            | 取消工作表或工作簿的保护。如果工作表或工作簿不是受保护的，则此方法不起作用。 |

**属性**

|                                                              | 名称                             | 说明                                                         |
| ------------------------------------------------------------ | -------------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**                  | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoScaling**                  | 如果 ET 对三维图表进行缩放，使之与等效的二维图表的大小相近，则该属性值为 **True**。**RightAngleAxes** 属性必须为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BackWall**                     | 返回 **Walls** 对象，该对象允许用户单独对三维图表的背景墙进行格式设置。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BarShape**                     | 返回或设置用于三维条形图或柱形图的形状。**XlBarShape** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ChartArea**                    | 返回一个 **ChartArea** 对象，该对象表示图表的整个图表区。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ChartStyle**                   | 返回或设置图表的图表样式。可读/写 **Variant** 类型。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ChartTitle**                   | 返回一个 **ChartTitle** 对象，该对象表示指定图表的标题。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ChartType**                    | 返回或设置图表类型。**XlChartType** 类型，可读写。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CodeName**                     | 返回对象的代码名。**String** 型，只读。                      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                      | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DataTable**                    | 返回一个 **DataTable** 对象，该对象表示图表的模拟运算表。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DepthPercent**                 | 返回或设置三维图表的深度，以图表宽度的百分比表示（有效范围从 20% 到 2000%）。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayBlanksAs**              | 返回或设置在图表中绘制空白单元格的方式。可以是 **XlDisplayBlanksAs** 常量之一。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Elevation**                    | 返回或设置三维图表视图的仰角（以角度为单位）。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Floor**                        | 返回一个 **Floor** 对象，该对象表示三维图表的基底。只读。返回值类型为Floor。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **GapDepth**                     | 以数据标志宽度的形式返回或设置三维图表中数据系列之间的距离，本属性的值必须在 0 和 500 之间。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HasAxis**                      | 返回或设置图表上显示的坐标轴。**Variant** 类型，可读写。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HasDataTable**                 | 如果图表有模拟运算表，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HasLegend**                    | 如果图表有图例，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HasTitle**                     | 如果坐标轴或图表有可见标题，则为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HeightPercent**                | 返回或设置三维图表的高度，以图表宽度的百分比表示（有效范围从 5% 到 500%）。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Hyperlinks**                   | 返回一个 **Hyperlinks** 集合，它代表图表的超链接。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Index**                        | 返回 **Long** 值，它代表对象在其同类对象所组成的集合内的索引号。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Legend**                       | 返回一个 **Legend** 对象，该对象表示指定图表的图例。只读。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MailEnvelope**                 | 代表文档的电子邮件标题。                                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**                         | 返回或设置一个 **String** 值，它代表对象的名称。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Next**                         | 返回代表下一个工作表的 **Worksheet** 对象。                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PageSetup**                    | 返回一个 **PageSetup** 对象，它包含用于指定对象的所有页面设置。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                       | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Perspective**                  | 返回或设置一个 **Long** 值，它代表三维图表视图的透视系数。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PivotLayout**                  | 返回一个 **PivotLayout** 对象，该对象表示数据透视表中字段的位置以及数据透视图中坐标轴的位置。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PlotArea**                     | 返回一个 **PlotArea** 对象，该对象表示图表的绘图区。只读。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PlotBy**                       | 返回或设置行或列在图表中作为数据系列使用的方式。可为以下 **XlRowCol** 常量之一：**xlColumns** 或 **xlRows**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PlotVisibleOnly**              | 如果仅绘制可见单元格，则该值为 **True**。如果可见单元格和隐藏单元格都绘制，则该值为 **False**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Previous**                     | 返回代表下一个工作表的 **Worksheet** 对象。                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PrintedCommentPages**          | 返回将为当前图表打印的批注页的数量。只读。返回类型类型为**Long。** |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ProtectContents**              | 如果工作表内容是受保护的，则为 **True**。对于图表，这样会保护整个图表。要打开内容保护，请使用 **Protect** 方法，并将 *Contents* 参数设置为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ProtectData**                  | 如果用户不能更改系列公式，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ProtectDrawingObjects**        | 如果形状是受保护的，则为 **True**。要打开形状保护，请使用 **Protect** 方法，并将 *DrawingObjects* 参数设置为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ProtectFormatting**            | 如果用户不能更改格式，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ProtectSelection**             | 如果不能选定图表元素，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ProtectionMode**               | 如果启用了用户界面专用保护，则为 **True**。要打开用户界面保护，请使用 **Protect** 方法，并将 *UserInterfaceOnly* 参数设置为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RightAngleAxes**               | 如果图表的坐标轴为直角，并与图表的转角或仰角无关，则该值为 **True**。仅应用于三维折线图、柱形图和条形图。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Rotation**                     | 以度为单位返回或设置三维图表视图的转角（绘图区绕 Z 轴的转角）。此属性的取值必须介于 0 到 360 之间，三维条形图除外（从 0 到 44 之间）。默认值是 20。仅适用于三维图表。**Variant** 型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shapes**                       | 返回一个 **Shapes** 集合，它代表图表工作表上所有的形状。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowAllFieldButtons**          | 返回或设置是否在数据透视图上显示所有字段按钮。可读/写。返回值类型为**Boolean。** |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowAxisFieldButtons**         | 返回或设置是否要在数据透视图上显示坐标轴字段按钮。可读/写。返回值类型为**Boolean。** |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowDataLabelsOverMaximum**    | 返回或设置一个布尔值，该值表示在数值大于数值轴上的最大值时是否显示数据标签。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowLegendFieldButtons**       | 返回或设置是否要在数据透视图上显示图例字段按钮。可读/写。返回值类型为**Boolean。** |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowReportFilterFieldButtons** | 返回或设置是否要在数据透视图上显示报表筛选字段按钮。可读/写。返回值类型为**Boolean。** |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowValueFieldButtons**        | 返回或设置是否要在数据透视图上显示值字段按钮。可读/写。返回值类型为**Boolean。** |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SideWall**                     | 返回 **Walls** 对象，该对象允许用户单独对三维图表的侧面墙进行格式设置。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Tab**                          | 为图表返回一个 **Tab** 对象。                                |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Visible**                      | 返回或设置一个 **XlSheetVisibility** 值，它确定对象是否可见。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Walls**                        | 返回一个 [**Walls**](https://qn.cache.wpscdn.cn/encs/doc/office_v13/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Walls/Walls%20%E5%AF%B9%E8%B1%A1.htm) 对象，该对象表示三维图表的背景墙。只读。 |

**成员方法**

#### **Chart.Activate**

使当前图表成为活动图表。

**语法**

**express.Activate()**

*express*   一个代表 **Chart** 对象的变量。

#### **Chart.ApplyChartTemplate**

将标准图表类型或自定义图表类型应用于图表。

**语法**

**express.ApplyChartTemplate()**

*express*   一个代表 **Chart** 对象的变量。

**说明**

此方法不支持采用原图表或组合图的常量。.

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**           |
| ---------- | ------------- | ------------ | ------------------ |
| *Filename* | 必选          | **String**   | 图表模板的文件名。 |

#### **Chart.ApplyDataLabels**

将数据标签应用到图表中的所有系列。

**语法**

**express.ApplyDataLabels(Type, LegendKey, AutoText, HasLeaderLines, ShowSeriesName, ShowCategoryName, ShowValue, ShowPercentage, ShowBubbleSize, Separator)**

*express*   一个代表 **Chart** 对象的变量。

**参数**

| **名称**           | **必选/可选** | **数据类型**         | **说明**                                                     |
| ------------------ | ------------- | -------------------- | ------------------------------------------------------------ |
| *Type*             | 可选          | **XlDataLabelsType** | 要应用的数据标签的类型。                                     |
| *LegendKey*        | 可选          | **Variant**          | 如果为 True，则在数据点旁边显示图例项标示。默认值为 False。  |
| *AutoText*         | 可选          | **Variant**          | 如果对象根据内容自动生成相应的文字，则为 True。              |
| *HasLeaderLines*   | 可选          | **Variant**          | 对于 Chart 和 Series 对象，如果数据系列有引导线，则为 True。 |
| *ShowSeriesName*   | 可选          | **Variant**          | 传递布尔值以启用或禁用数据标签的系列名称。                   |
| *ShowCategoryName* | 可选          | **Variant**          | 传递布尔值以启用或禁用数据标签的分类名称。                   |
| *ShowValue*        | 可选          | **Variant**          | 传递布尔值以启用或禁用数据标签的值。                         |
| *ShowPercentage*   | 可选          | **Variant**          | 传递布尔值以启用或禁用数据标签的百分比。                     |
| *ShowBubbleSize*   | 可选          | **Variant**          | 传递布尔值以启用或禁用数据标签的气泡大小。                   |
| *Separator*        | 可选          | **Variant**          | 数据标签的分隔符。                                           |

**说明**

**注释：**When using the **Separator** parameter.
If you use a string, you will get a string as the separator. If you use **xlDataLabelSeparatorDefault** (= 1), you will get the default data label separator, which is either a comma or a newline, depending on the data label.
When a value of "1" is returned, it indicates that the user has not changed the default separator which is a comma ",". You can also pass a value of "1" to change the separator back to the default separator.
The chart must first be active before you can access the data labels programmatically; otherwise a run-time error occurs.

| **XlDataLabelsType** 可为以下 **XlDataLabelsType** 常量之一。 |
| ------------------------------------------------------------ |
| **xlDataLabelsShowBubbleSizes**。数据标签的气泡尺寸。        |
| **xlDataLabelsShowLabelAndPercent**。占总数的百分比及数据点所属的分类。仅用于饼图和圆环图。 |
| **xlDataLabelsShowPercent**。占总数的百分比。仅用于饼图和圆环图。 |
| **xlDataLabelsShowLabel**。数据点所属的分类。                |
| **xlDataLabelsShowNone**。无数据标签。                       |
| **xlDataLabelsShowValue**。默认。数据点的值（若未指定此参数，则假定为此值）。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对 Chart1 上的第一个数据系列应用分类标签。*/ Application.Charts.Item("Chart1").SeriesCollection(1).ApplyDataLabels(xlDataLabelsShowLabel)` |

#### **Chart.ApplyLayout**

应用功能区中显示的版式。

**语法**

**express.ApplyLayout(Layout, ChartType)**

*express*   一个代表 **Chart** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型**    | **说明**                                      |
| ----------- | ------------- | --------------- | --------------------------------------------- |
| *Layout*    | 必选          | **Long**        | 指定版式类型。版式类型由 1 到 10 的数字表示。 |
| *ChartType* | 可选          | **XlChartType** | 图表的类型。                                  |

**说明**

在当前图表类型上使用版式时，会将介于 1 到 10 之间的数字应用于该图表类型。还可以将某一图表类型的版式应用于另一个图表类型。例如，可以将折线图中的可用版式应用于柱形图。版式只会添加可用于该特定图表类型的图表元素。

#### **Chart.Axes**

返回一个代表图表上单个坐标轴或坐标轴集合的某个对象。

**语法**

**express.Axes(Type, AxisGroup)**

*express*   一个代表 **Chart** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型**    | **说明**                                                     |
| ----------- | ------------- | --------------- | ------------------------------------------------------------ |
| *Type*      | 可选          | **Variant**     | 指定要返回的坐标轴。可为以下 XlAxisType 常量之一：xlValue、xlCategory 或 xlSeriesAxis（xlSeriesAxis 仅对三维图表有效）。 |
| *AxisGroup* | 可选          | **XlAxisGroup** | 指定坐标轴组。如果省略该参数，则使用主坐标轴组。三维图表仅有一个坐标轴组。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为 Chart1 的分类轴添加轴标签。*/ function test() {     let axes = Application.Charts.Item("Chart1").Axes(xlCategory)     axes.HasTitle = true     axes.AxisTitle.Text = "July Sales" }  /*本示例将关闭 Chart1 中分类轴的主要网格线。*/ Application.Charts.Item("Chart1").Axes(xlCategory).HasMajorGridlines = false  /*本示例将关闭 Chart1 中所有坐标轴的网格线。*/ function test() {     for(let i = 1; i <= Application.Charts.Item("Chart1").Axes().Count; i++ ){         Application.Charts.Item("Chart1").Axes().Item(i).HasMajorGridlines = false         Application.Charts.Item("Chart1").Axes().Item(i).HasMinorGridlines = false     } }` |

#### **Chart.ChartGroups**

返回一个对象，该对象表示图表中单个图表组（**ChartGroup** 对象）或所有图表组的集合（**ChartGroups** 对象）。返回的集合中包括每种类型的图表组。

**语法**

**express.ChartGroups(Index)**

*express*   一个代表 **Chart** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**     |
| -------- | ------------- | ------------ | ------------ |
| *Index*  | 可选          | **Variant**  | 图表组编号。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示 Chart1 中第一个图表组的涨跌柱线，并对其颜色进行设置。本示例应在包含两个有一个或多个交叉数据点的系列的二维折线图上运行。*/ function test() {     let chart = Application.Charts.Item("Chart1").ChartGroups(1)     chart.HasUpDownBars = true     chart.DownBars.Interior.ColorIndex = 3     chart.UpBars.Interior.ColorIndex = 5 }` |

#### **Chart.ChartObjects**

返回一个对象，它代表工作表上的一个嵌入式图表（**ChartObject** 对象）或所有嵌入式图表的集合（**ChartObjects** 对象）。

**语法**

**express.ChartObjects(Index)**

*express*   一个代表 **Chart** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**      |
| -------- | ------------- | ------------ | ------------- |
| *Index*  | 可选          | **Variant**  | {description} |

**说明**

该方法不等同于 **Charts** 属性，它返回嵌入式图表，而 **Charts** 属性返回图表工作表。使用 **Chart** 属性可返回嵌入式图表的 **Chart** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向工作表 Sheet1 上的第一个嵌入式图表中添加标题。*/ function test() {     let chart = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart     chart.HasTitle = true     chart.ChartTitle.Text = "1995 Rainfall Totals by Month" }  /*本示例在 Sheet1 上的第一个嵌入式图表中新建一个系列，新系列的数据源为 Sheet1 的 B1:B10 区域。*/ function test() {     Application.Worksheets.Item("Sheet1").ChartObjects(1).Activate()     Application.ActiveChart.SeriesCollection.Add(Worksheets.Item("Sheet1").Range("B1:B10")) }  /*本示例清除工作表 Sheet1 上第一个嵌入式图表的格式设置。*/ Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartArea.ClearFormats()` |

#### **Chart.ChartWizard**

修改给定图表的属性。可使用本方法快速设置图表的格式，而不必逐个设置所有属性。本方法是非交互式的，并且仅更改指定的属性。

**语法**

**express.ChartWizard(Source, Gallery, Format, PlotBy, CategoryLabels, SeriesLabels, HasLegend, Title, CategoryTitle, ValueTitle, ExtraTitle)**

*express*   一个代表 **Chart** 对象的变量。

**参数**

| **名称**         | **必选/可选** | **数据类型** | **说明**                                                     |
| ---------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Source*         | 可选          | **Variant**  | 包含新图表源数据的区域。如果省略本参数，ET 将编辑活动图表工作表或活动工作表上处于选定状态的图表。 |
| *Gallery*        | 可选          | **Variant**  | 用于指定图表类型的 XlChartType 的常量之一。                  |
| *Format*         | 可选          | **Variant**  | 内置自动套用格式的选项编号。可为从 1 到 10 的数字，其取值取决于库的类型。如果省略此参数，ET 将根据库的类型和数据源选择默认值。 |
| *PlotBy*         | 可选          | **Variant**  | 指定每个系列的数据是来自行还是来自列。可以是以下 XlRowCol 常量之一：xlRows 或 xlColumns。 |
| *CategoryLabels* | 可选          | **Variant**  | 指定包含分类标签的源范围内的行数或列数的整数。合法值为从 0（零）至小于相应分类或系列的最大个数间的某一数字。 |
| *SeriesLabels*   | 可选          | **Variant**  | 指定包含系列标志的源范围内的行数或列数的整数。合法值为从 0（零）至小于相应分类或系列的最大个数间的某一数字。 |
| *HasLegend*      | 可选          | **Variant**  | 若要包括图例，则为 True。                                    |
| *Title*          | 可选          | **Variant**  | 图表标题文字。                                               |
| *CategoryTitle*  | 可选          | **Variant**  | 分类轴标题文字。                                             |
| *ValueTitle*     | 可选          | **Variant**  | 数值轴标题文字。                                             |
| *ExtraTitle*     | 可选          | **Variant**  | 三维图表的系列轴标题，或二维图表的次数值轴标题。             |

**说明**

如果省略 *Source*，并且选定内容不是活动工作表中的嵌入图表或者活动工作表不是当前存在的图表，则该方法失效并产生错误。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例重新设置 Chart1 的格式，将其改为折线图，添加图例，并添加分类轴标题和数值轴标题。*/ Application.Charts.Item("Chart1").ChartWizard(undefined, xlLine, undefined, undefined, undefined, undefined, true, undefined, "Year", "Sales") ` |

#### **Chart.CheckSpelling**

检查对象的拼写。

**语法**

**express.CheckSpelling(CustomDictionary, IgnoreUppercase, AlwaysSuggest, SpellLang)**

*express*   一个代表 **Chart** 对象的变量。

**参数**

| **名称**           | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------------ | ------------- | ------------ | ------------------------------------------------------------ |
| *CustomDictionary* | 可选          | **Variant**  | 一个字符串，它表示自定义词典的文件名，如果在主词典中找不到单词，则会到此词典中查找。如果省略此参数，则使用当前指定的词典。 |
| *IgnoreUppercase*  | 可选          | **Variant**  | 如果为 True，则 ET 忽略所有字母都是大写的单词。如果为 False，则 ET 检查所有字母都是大写的单词。如果省略此参数，将使用当前的设置。 |
| *AlwaysSuggest*    | 可选          | **Variant**  | 如果为 True，则 ET 在找到不正确拼写时显示建议的替换拼写列表。如果为 False，ET 将等待输入正确的拼写。如果省略此参数，将使用当前的设置。 |
| *SpellLang*        | 可选          | **Variant**  | 当前所用词典的语言。它可以是由 LanguageID 属性使用的 MsoLanguageID 值之一。 |

**说明**

此方法没有返回值；ET 显示**“拼写检查”**对话框。

要检查工作表中的页眉、页脚和对象，可对 **Worksheet** 对象使用此方法。

#### **Chart.ClearToMatchStyle**

清除图表元素格式以改为自动格式。

**语法**

**express.ClearToMatchStyle()**

*express*   一个代表 **Chart** 对象的变量。

**说明**

使用此方法将图表元素格式重置为自动格式。如果在图表上使用此方法，则将重置所有格式（包括替代项）。

#### **Chart.Copy**

将工作表复制到工作簿的另一位置。

**语法**

**express.Copy(Before, After)**

*express*   一个代表 **Chart** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Before* | 可选          | **Variant**  | 将要在其之前放置所复制工作表的工作表。如果指定了 After，则不能指定 Before。 |
| *After*  | 可选          | **Variant**  | 将要在其之后放置所复制工作表的工作表。如果指定了 Before，则不能指定 After。 |

**说明**

如果既不指定 *Before* 也不指定 *After*，则 ET 将新建一个工作簿，其中包含复制的工作表。

**Copy** 方法不支持图表对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例复制工作表 Sheet1，并将其放置在工作表 Sheet3 之后。*/ Application.Worksheets.Item("Sheet1").Copy(null,Worksheets.Item("Sheet3"))` |

#### **Chart.CopyPicture**

将所选对象作为图片复制到剪贴板。

**语法**

**express.CopyPicture(Appearance, Format, Size)**

*express*   一个代表 **Chart** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型**            | **说明**                                                     |
| ------------ | ------------- | ----------------------- | ------------------------------------------------------------ |
| *Appearance* | 可选          | **XlPictureAppearance** | 指定图片的复制方式。默认值为 xlScreen。                      |
| *Format*     | 可选          | **XlCopyPictureFormat** | 图片的格式。默认值为 xlPicture。                             |
| *Size*       | 可选          | **XlPictureAppearance** | 当对象是图表工作表中的图表（不是工作表中的嵌入式图表）时，此参数代表复制图片的大小。默认值为 xlPrinter。 |

#### **Chart.Delete**

删除对象。

**语法**

**express.Delete()**

*express*   一个代表 **Chart** 对象的变量。

#### **Chart.Evaluate**

将一个 ET 名称转换为一个对象或者一个值。

**语法**

**express.Evaluate(Name)**

*express*   一个代表 **Chart** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                     |
| -------- | ------------- | ------------ | ---------------------------- |
| *Name*   | 必选          | **Variant**  | 使用 ET 命名约定的对象名称。 |

**说明**

该方法可使用下列 ET 名称类型：

- A1 格式引用。可以通过 A1 格式表示法引用单个单元格。所有引用均视为绝对引用。
- 区域。在引用中可以使用区域、交集和联合运算符（分别为冒号、空格和逗号）。
- 定义的名称。可用宏语言指定任何名称。
- 外部引用。可以使用 ! 运算符引用另一工作簿中的单元格或已定义的名称，例如，`Evaluate("[BOOK1.XLS]Sheet1!A1")`。
- 图表对象。可以指定任何图表对象名称（如“Legend”、“Plot Area”或“Series 1”），以访问该对象的属性和方法。例如，`Charts("Chart1").Evaluate("Legend").Font.Name` 返回图例中所用字体的名称。

**注释**：使用方括号（例如，“[A1:C5]”，JsApi不支持方括号的用法）、或直接调用对应对象与用字符串参数调用 **Evaluate** 方法是等效的。例如，下列表达式对是等效的。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Rnage("a1").Value2 = 25 Application.Evaluate("A1").Value2 = 25 		 trigVariable = SIN(45) trigVariable = Application.Evaluate("SIN(45)") 		 Set firstCellInSheet = Application.Workbooks.Item("BOOK1.XLS").Sheets.Item(4).Range("A1") Set firstCellInSheet = Application.Workbooks.Item("BOOK1.XLS").Sheets.Item(4).Evaluate("A1")` |

使用方括号的优点在于代码较短。使用 **Evaluate** 的优点在于参数是字符串，这样您既可以在代码中构造该字符串，也可以使用 Visual Basic 变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例将工作表 Sheet1 上 A1 单元格的字体设置为加粗。*/ function test() {     Application.Worksheets.Item("Sheet1").Activate()     boldCell = "A1"     Application.Evaluate(boldCell).Font.Bold = true }` |

#### **Chart.Export**

以图形格式导出图表。

**语法**

**express.Export(Filename, FilterName, Interactive)**

*express*   一个代表 **Chart** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Filename*    | 必选          | **String**   | 被导出的文件的名称。                                         |
| *FilterName*  | 可选          | **Variant**  | 图形过滤器出现在注册表中时独立于语言的名称。                 |
| *Interactive* | 可选          | **Variant**  | 如果为 True，则显示包含筛选器特定选项的对话框。如果为 False，则 ET 使用筛选器的默认值。默认值是 False。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例以 GIF 文件格式导出第一个图表。*/ Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Export("current_sales.gif", "GIF")` |

#### **Chart.ExportAsFixedFormat**

导出为指定格式的文件。

**语法**

**express.ExportAsFixedFormat(Type, Filename, Quality, IncludeDocProperties, IgnorePrintAreas, From, To, OpenAfterPublish, FixedFormatExtClassPtr)**

*express*   一个代表 **Chart** 对象的变量。

**参数**

| **名称**                 | **必选/可选** | **数据类型**          | **说明**                                                     |
| ------------------------ | ------------- | --------------------- | ------------------------------------------------------------ |
| *Type*                   | 必选          | **XlFixedFormatType** | 要导出为的文件格式类型。                                     |
| *Filename*               | 可选          | **Variant**           | 要保存的文件的文件名。可以包括完整路径，否则 ET 会将文件保存在当前文件夹中。 |
| *Quality*                | 可选          | **Variant**           | 可选 XlFixedFormatQuality。指定已发布文件的质量。            |
| *IncludeDocProperties*   | 可选          | **Variant**           | 若要包括文档属性，则为 True；否则为 False。                  |
| *IgnorePrintAreas*       | 可选          | **Variant**           | 若要忽略发布时设置的任何打印区域，则为 True；否则为 False。  |
| *From*                   | 可选          | **Variant**           | 发布的起始页码。如果省略此参数，则从起始位置开始发布。       |
| *To*                     | 可选          | **Variant**           | 发布的终止页码。如果省略此参数，则发布至最后一页             |
| *OpenAfterPublish*       | 可选          | **Variant**           | 若要在发布文件后在查看器中显示文件，则为 True；否则为 False。 |
| *FixedFormatExtClassPtr* | 可选          | **Variant**           | 指向 FixedFormatExt 类的指针。                               |

**说明**

此方法还支持初始化加载项，以将文件导出为固定格式的文件。例如，如果存在转换器，则 ET 将执行文件格式转换。转换通常由用户执行。

#### **Chart.GetChartElement**

返回指定的 X 坐标和 Y 坐标上图表元素的信息。本方法稍有与众不同之处：调用时只须指定前两个参数，在本方法执行期间，ET 为其余参数赋值，本方法返回后应检验这些参数的值。

**语法**

**express.GetChartElement(x, y, ElementID, Arg1, Arg2)**

*express*   一个代表 **Chart** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *x*         | 必选          | **Long**     | 图表元素的 X 坐标。                                          |
| *y*         | 必选          | **Long**     | 图表元素的 Y 坐标。                                          |
| *ElementID* | 必选          | **Long**     | 该方法返回时，此参数包含指定坐标的图表元素的 XLChartItem 值。有关详细信息，请参阅“备注”部分。 |
| *Arg1*      | 必选          | **Long**     | 该方法返回时，该参数包含与图表元素相关的信息。有关详细信息，请参阅“备注”部分。 |
| *Arg2*      | 必选          | **Long**     | 该方法返回时，该参数包含与图表元素相关的信息。有关详细信息，请参阅“备注”部分。 |

**说明**

本方法返回的 *ElementID* 值决定了 *Arg1* 和 *Arg2* 所包含的信息，如下表所示。

| ElementID 常量              | 常量值 | Arg1         | Arg2            |
| --------------------------- | ------ | ------------ | --------------- |
| **xlAxis**                  | 21     | AxisIndex    | AxisType        |
| **xlAxisTitle**             | 17     | AxisIndex    | AxisType        |
| **xlDisplayUnitLabel**      | 30     | AxisIndex    | AxisType        |
| **xlMajorGridlines**        | 15     | AxisIndex    | AxisType        |
| **xlMinorGridlines**        | 16     | AxisIndex    | AxisType        |
| **xlPivotChartDropZone**    | 32     | DropZoneType | 无              |
| **xlPivotChartFieldButton** | 31     | DropZoneType | PivotFieldIndex |
| **xlDownBars**              | 20     | GroupIndex   | 无              |
| **xlDropLines**             | 26     | GroupIndex   | 无              |
| **xlHiLoLines**             | 25     | GroupIndex   | 无              |
| **xlRadarAxisLabels**       | 27     | GroupIndex   | 无              |
| **xlSeriesLines**           | 22     | GroupIndex   | 无              |
| **xlUpBars**                | 18     | GroupIndex   | 无              |
| **xlChartArea**             | 2      | 无           | 无              |
| **xlChartTitle**            | 4      | 无           | 无              |
| **xlCorners**               | 6      | 无           | 无              |
| **xlDataTable**             | 7      | 无           | 无              |
| **xlFloor**                 | 23     | 无           | 无              |
| **xlLeaderLines**           | 29     | 无           | 无              |
| **xlLegend**                | 24     | 无           | 无              |
| **xlNothing**               | 28     | 无           | 无              |
| **xlPlotArea**              | 19     | 无           | 无              |
| **xlWalls**                 | 5      | 无           | 无              |
| **xlDataLabel**             | 7      | SeriesIndex  | PointIndex      |
| **xlErrorBars**             | 9      | SeriesIndex  | 无              |
| **xlLegendEntry**           | 12     | SeriesIndex  | 无              |
| **xlLegendKey**             | 13     | SeriesIndex  | 无              |
| **xlSeries**                | 3      | SeriesIndex  | PointIndex      |
| **xlShape**                 | 14     | ShapeIndex   | 无              |
| **xlTrendline**             | 8      | SeriesIndex  | TrendLineIndex  |
| **xlXErrorBars**            | 10     | SeriesIndex  | 无              |
| **xlYErrorBars**            | 11     | SeriesIndex  | 无              |

 

下表说明了本方法返回后，参数 *Arg1* 和 *Arg2* 的含义。

| 参数            | 说明                                                         |
| --------------- | ------------------------------------------------------------ |
| AxisIndex       | 指定坐标轴是主坐标轴还是次坐标轴。可为以下 **XlAxisGroup** 常量之一：**xlPrimary** 或 **xlSecondary**。 |
| AxisType        | 指定坐标轴类型。可为以下 **XlAxisType** 常量之一：**xlCategory**、**xlSeriesAxis** 或 **xlValue**。 |
| DropZoneType    | 指定拖放区的类型：列、数据、页或行字段。可为以下 **XlPivotFieldOrientation** 常量之一：**xlColumnField**、**xlDataField**、**xlPageField** 或 **xlRowField**。列和行字段常量分别指定了系列和分类字段。 |
| GroupIndex      | 指定特定图表组在 **ChartGroups** 集合内的偏移量。            |
| PivotFieldIndex | 指定特定列（数据系列）、数据、页或行（类别）字段在 **PivotFields** 集合中的偏移量。如果拖放区域类型为 **xlDataField**，则返回 -1。 |
| PointIndex      | 指定系列中的特定点在 **Points** 集合中的偏移量。-1 表示选定了所有数据点。 |
| SeriesIndex     | 指定特定系列在 **Series** 集合中的偏移量。                   |
| ShapeIndex      | 指定特定形状在 **Shapes** 集合中的偏移量。                   |
| TrendlineIndex  | 指定某个系列中特定趋势线在 **Trendlines** 集合中的偏移量。   |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*当鼠标移动到图表图例上时，本示例发出警告。*/ function test() {     let IDNum      let a     let b     Application.ActiveChart.GetChartElement(X, Y, IDNum, a, b)     if(IDNum == xlLegendEntry) {         alert("WARNING: Move away from the legend")     } }` |

#### **Chart.Location**

将图表移动到新位置。

**语法**

**express.Location(Where, Name)**

*express*   一个代表 **Chart** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**        | **说明**                                                     |
| -------- | ------------- | ------------------- | ------------------------------------------------------------ |
| *Where*  | 必选          | **XlChartLocation** | 图表移动的目标位置。                                         |
| *Name*   | 可选          | **Variant**         | 如果 Where 为 xlLocationAsObject，则该参数为必选参数。如果 Where 为 xlLocationAsObject，则该参数为嵌入该图表的工作表的名称。如果 Where 为 xlLocationAsNewSheet，则该参数为新工作表的名称。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将嵌入图表移至新图表工作表“Monthly Sales”。*/ Application.Worksheets.Item(1).ChartObjects(1).Chart.Location(xlLocationAsNewSheet, "Monthly Sales")` |

#### **Chart.Move**

将图表移到工作簿的另一位置。

**语法**

**express.Move(Before, After)**

*express*   一个代表 **Chart** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Before* | 可选          | **Variant**  | 将要在其之前放置所移动图表的工作表。如果指定了 After，则不能指定 Before。 |
| *After*  | 可选          | **Variant**  | 将要在其之后放置所移动图表的工作表。如果指定了 Before，则不能指定 After。 |

**说明**

如果既不指定 *Before* 也不指定 *After*，则 ET 将新建一个工作簿并将要移动的图表移到新工作簿中。

#### **Chart.OLEObjects**

返回一个对象，它代表图表或工作表上的单个 OLE 对象 (**OLEObject** ) 或所有 OLE 对象的集合（**OLEObjects** 集合）。只读。

**语法**

**express.OLEObjects(Index)**

*express*   一个代表 **Chart** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**               |
| -------- | ------------- | ------------ | ---------------------- |
| *Index*  | 可选          | **Variant**  | OLE 对象的名称或编号。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例创建工作表 Sheet1 上 OLE 对象的链接类型列表。该列表将出现在此示例新建的工作表中。*/ function test() {     let newSheet = Application.Worksheets.Add()     let j = 2     newSheet.Range("A1").Value2 = "Name"     newSheet.Range("B1").Value2 = "Link Type"     for(let i = 1; i <= Application.Worksheets.Item("Sheet1").OLEObjects().Count; i++){         newSheet.Cells.Item(j, 1).Value2 = Application.Worksheets.Item("Sheet1").OLEObjects().Item(i).Name         if(Application.Worksheets.Item("Sheet1").OLEObjects().Item(i).OLEType == xlOLELink){             newSheet.Cells.Item(j, 2) = "Linked"         }         else{             newSheet.Cells.Item(j, 2) = "Embedded"         }         j++     } }` |

#### **Chart.Paste**

将剪贴板中的图表数据粘贴到指定的图表中。

**语法**

**express.Paste(Type)**

*express*   一个代表 **Chart** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Type*   | 可选          | **Variant**  | 指定要粘贴的图表信息（如果剪贴板中有一个图表）。可以为以下 XlPasteType 常量之一：xlPasteFormats、xlPasteFormulas 或 xlPasteAll。默认值为 xlPasteAll。如果剪贴板上的数据不是图表数据，则不能使用该参数。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将工作表 Sheet1 上单元格区域 B1:B5 中的数据粘贴到 Chart1 中。*/ function test() {     Application.Worksheets.Item("Sheet1").Range("B1:B5").Copy()     Application.Charts.Item("Chart1").Paste() }` |

#### **Chart.PrintOut**

打印对象。

**语法**

**express.PrintOut(From, To, Copies, Preview, ActivePrinter, PrintToFile, Collate, PrToFileName)**

*express*   一个代表 **Chart** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型** | **说明**                                                     |
| --------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *From*          | 可选          | **Variant**  | 打印的开始页号。如果省略此参数，则从起始位置开始打印。       |
| *To*            | 可选          | **Variant**  | 打印的终止页号。如果省略此参数，则打印至最后一页。           |
| *Copies*        | 可选          | **Variant**  | 打印份数。如果省略此参数，则只打印一份。                     |
| *Preview*       | 可选          | **Variant**  | 如果为 True，ET 将在打印对象之前调用打印预览。如果为 False（或省略该参数），则立即打印对象。 |
| *ActivePrinter* | 可选          | **Variant**  | 设置活动打印机的名称。                                       |
| *PrintToFile*   | 可选          | **Variant**  | 如果为 True，则打印到文件。如果没有指定 PrToFileName，ET 将提示用户输入要使用的输出文件的文件名。 |
| *Collate*       | 可选          | **Variant**  | 如果为 True，则逐份打印多个副本。                            |
| *PrToFileName*  | 可选          | **Variant**  | 如果 PrintToFile 设为 True，则该参数指定要打印到的文件名。   |

**说明**

*From* 和 *To* 所描述的“页”指的是要打印的页，并非指定工作表或工作簿中的全部页。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例打印当前活动工作表。*/ Application.ActiveSheet.PrintOut()` |

#### **Chart.PrintPreview**

按对象打印后的外观效果显示对象的预览。

**语法**

**express.PrintPreview(EnableChanges)**

*express*   一个代表 **Chart** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型** | **说明**                                                     |
| --------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *EnableChanges* | 可选          | **Variant**  | 传递 Boolean 值，以指定用户是否可更改边距和打印预览中可用的其他页面设置选项。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例在打印预览中显示工作表 Sheet1。*/ Application.Worksheets.Item("Sheet1").PrintPreview()` |

#### **Chart.Protect**

保护图表使其不被修改。

**语法**

**express.Protect(Password, DrawingObjects, Contents, Scenarios, UserInterfaceOnly)**

*express*   一个代表 **Chart** 对象的变量。

**参数**

| **名称**            | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Password*          | 可选          | **Variant**  | 一个字符串，该字符串为工作表或工作簿指定区分大小写的密码。如果省略此参数，不用密码就可以取消对工作表或工作簿的保护。否则，必须指定密码才能取消对工作表或工作簿的保护。如果忘记了密码，就无法取消对工作表或工作簿的保护。 |
| *DrawingObjects*    | 可选          | **Variant**  | 如果为 True，则保护形状。默认值是 True。                     |
| *Contents*          | 可选          | **Variant**  | 如果为 True，则保护内容。对于图表，这样会保护整个图表。对于工作表，这样会保护锁定的单元格。默认值是 True。 |
| *Scenarios*         | 可选          | **Variant**  | 如果为 True，则保护方案。此参数仅对工作表有效。默认值是 True。 |
| *UserInterfaceOnly* | 可选          | **Variant**  | 如果为 True，则保护用户界面，但不保护宏。如果省略此参数，则既保护宏也保护用户界面。 |

#### **Chart.Refresh**

立即重新绘制指定的图表。

**语法**

**express.Refresh()**

*express*   一个代表 **Chart** 对象的变量。

#### **Chart.SaveAs**

将对图表或工作表的更改保存到另一不同文件中。

**语法**

**express.SaveAs(Filename, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, AddToMru, TextCodepage, TextVisualLayout, Local)**

*express*   一个代表 **Chart** 对象的变量。

**参数**

| **名称**              | **必选/可选** | **数据类型** | **说明**                                                     |
| --------------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Filename*            | 必选          | **String**   | Variant。一个字符串，表示要保存的文件名。可包含完整路径。如果不指定路径，ET 将文件保存到当前文件夹中。 |
| *FileFormat*          | 可选          | **Variant**  | 保存文件时使用的文件格式。要查看有效的选项列表，请参阅 FileFormat 属性。对于现有文件，默认采用上一次指定的文件格式；对于新文件，默认采用当前所用 ET 版本的格式。 |
| *Password*            | 可选          | **Variant**  | 它是一个区分大小写的字符串（最长不超过 15 个字符），用于指定文件的保护密码。 |
| *WriteResPassword*    | 可选          | **Variant**  | 一个表示文件写保护密码的字符串。如果文件保存时带有密码，但打开文件时不输入密码，则该文件以只读方式打开。 |
| *ReadOnlyRecommended* | 可选          | **Variant**  | 如果为 True，则在打开文件时显示一条消息，提示该文件以只读方式打开。 |
| *CreateBackup*        | 可选          | **Variant**  | 如果为 True，则创建备份文件。                                |
| *AddToMru*            | 可选          | **Variant**  | 如果为 True，则将该工作簿添加到最近使用的文件列表中。默认值是 False。 |
| *TextCodepage*        | 可选          | **Variant**  | 不在美国英语版的 ET 中使用。                                 |
| *TextVisualLayout*    | 可选          | **Variant**  | 不在美国英语版的 ET 中使用。                                 |
| *Local*               | 可选          | **Variant**  | 如果为 True，则以 ET（包括控制面板设置）的语言保存文件。如果为 False（默认值），则以 示例代码 (VBA) 的语言保存文件，其中 示例代码 (VBA) 通常为美国英语版本，除非从中运行 Workbooks.Open 的 VBA 项目是旧的已国际化的 XL5/95 VBA 项目。 |

**说明**

请使用同时包含大小写字母、数字和符号的强密码。弱密码不混合使用这些元素。强密码：Y6dh!et5。弱密码：House27。请使用您可以记住的强密码，这样就不必将它写下来。

#### **Chart.SaveChartTemplate**

向可用图表模板的列表添加自定义图表模板。

**语法**

**express.SaveChartTemplate(Filename)**

*express*   一个代表 **Chart** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**         |
| ---------- | ------------- | ------------ | ---------------- |
| *Filename* | 必选          | **String**   | 图表模板的名称。 |

**说明**

默认情况下，此方法将活动图表保存到用户的图表模板库中。如果指定了 UNC 或 URL，则图表将保存到指定位置。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例基于活动图表添加一个新图表模板。*/ Application.ActiveChart.SaveChartTemplate("Presentation Chart")` |

#### **Chart.Select**

选择对象。

**语法**

**express.Select(Replace)**

*express*   一个代表 **Chart** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型** | **说明**                                                     |
| --------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Replace* | 可选          | **Variant**  | （仅用于工作表）。如果为 True，则用指定的对象替换当前所选内容。如果为 False，则扩展当前所选内容以包括以前选择的对象和指定的对象。 |

**说明**

要选择单元格或单元格区域，请使用 **Select** 方法。要使单个单元格成为活动单元格，请使用 **Activate** 方法。

#### **Chart.SeriesCollection**

返回一个对象，它代表图表或图表组中的一个系列（**Series** 对象）或所有系列的集合（**SeriesCollection** 集合）。

**语法**

**express.SeriesCollection(Index)**

*express*   一个代表 **Chart** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**               |
| -------- | ------------- | ------------ | ---------------------- |
| *Index*  | 可选          | **Variant**  | 数据系列的名称或编号。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例打开 Chart1 中第一个数据系列的数据标签。*/ Application.Charts.Item("Chart1").SeriesCollection(1).HasDataLabels = true` |

#### **Chart.SetBackgroundPicture**

为图表设置背景图形。

**语法**

**express.SetBackgroundPicture(Filename)**

*express*   一个代表 **Chart** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**     |
| ---------- | ------------- | ------------ | ------------ |
| *Filename* | 必选          | **String**   | 图形文件名。 |

#### **Chart.SetDefaultChart**

指定 ET 新建图表时使用的图表模板的名称。

**语法**

**express.SetDefaultChart(Name)**

*express*   一个代表 **Chart** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Name*   | 必选          | **Variant**  | 指定新建图表时将使用的默认图表模板的名称。此名称可以是字符串，指定库中图表，进而指定用户定义的模板，也可以是一个特殊常量 xlBuiltIn，用于指定内置图表模板。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例将默认图表模板设置为名为“Monthly Sales”的自定义图表。*/ Application.ActiveChart.SetDefaultChart("Monthly Sales")` |

#### **Chart.SetElement**

设置图表上的图表元素。可读/写 **MsoChartElementType** 类型。

**语法**

**express.SetElement(Element)**

*express*   一个代表 **Chart** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型**            | **说明**           |
| --------- | ------------- | ----------------------- | ------------------ |
| *Element* | 必选          | **MsoChartElementType** | 指定图表元素类型。 |

**说明**

对于图表，**“布局”**选项卡中的以下命令对应于 **SetElement** 方法：

- **“标签”**组中的全部内容。
- **“坐标轴”**组中的全部内容。
- **“分析”**组中的全部内容。
- **“绘图区”**、**“图表背景墙”**和**“图表基底”**按钮。

**MsoChartElementType** 是常量枚举，引用上述所有命令。

 

 

 

 

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例使用不同常量值向活动图表设置图表元素。*/ function test() {     Application.ActiveChart.Axes(xlValue).MajorGridlines.Select()     Application.ActiveChart.SetElement(msoElementChartTitleCenteredOverlay)     Application.ActiveChart.SetElement(msoElementPrimaryCategoryGridLinesMinor)     Application.ActiveChart.Walls.Select()     Application.Application.CommandBars.Item("Clip Art").Visible = false     Application.ActiveChart.SetElement(msoElementChartFloorShow) }` |

#### **Chart.SetSourceData**

为指定图表设置源数据区域。

**语法**

**express.SetSourceData(Source, PlotBy)**

*express*   一个代表 **Chart** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Source* | 必选          | **Range**    | 包含源数据的区域。                                           |
| *PlotBy* | 可选          | **Variant**  | 指定数据绘制方式。可为以下 XlRowCol 常量之一：xlColumns 或 xlRows。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为第一个图表设置源数据区域。*/ Application.Charts.Item(1).SetSourceData(Sheets.Item(1).Range("a1:a10"), xlColumns)` |

#### **Chart.Unprotect**

取消工作表或工作簿的保护。如果工作表或工作簿不是受保护的，则此方法不起作用。

**语法**

**express.Unprotect(Password)**

*express*   一个代表 **Chart** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**                                                     |
| ---------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Password* | 可选          | **Variant**  | 指定用于解除图表保护的密码，此密码是区分大小写的。如果图表不设密码保护，则忽略该参数。 |

**说明**

如果您忘记了密码，将不能取消工作表或工作簿的保护。建议将密码和对应文档名妥善保存。

**成员属性**

#### **Chart.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **Chart** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() {     if(Application.ActiveWorkbook.Application.Value == "ET"){         alert("This is an ET Application object.")     }     else{         alert("This is not an ET Application object.")     } }` |

#### **Chart.AutoScaling**

如果 ET 对三维图表进行缩放，使之与等效的二维图表的大小相近，则该属性值为 **True**。**RightAngleAxes** 属性必须为 **True**。**Boolean** 类型，可读写。

**语法**

**express.AutoScaling**

*express*   一个代表 **Chart** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例自动对 Chart1 进行缩放。本示例应在三维图表上运行。*/ function test() {     let charts = Application.Charts.Item("Chart1")     charts.RightAngleAxes = true     charts.AutoScaling = true }` |

#### **Chart.BackWall**

返回 **Walls** 对象，该对象允许用户单独对三维图表的背景墙进行格式设置。只读。

**语法**

**express.BackWall**

*express*   一个代表 **Chart** 对象的变量。

#### **Chart.BarShape**

返回或设置用于三维条形图或柱形图的形状。**XlBarShape** 类型，可读写。

**语法**

**express.BarShape**

*express*   一个代表 **Chart** 对象的变量。

#### **Chart.ChartArea**

返回一个 **ChartArea** 对象，该对象表示图表的整个图表区。只读。

**语法**

**express.ChartArea**

*express*   一个代表 **Chart** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 Chart1 的图表区域内部颜色设置为红色，并将其边框颜色设置为蓝色。*/ function test() {     let area = Application.Charts.Item("Chart1").ChartArea     area.Interior.ColorIndex = 3     area.Border.ColorIndex = 5 }` |

#### **Chart.ChartStyle**

返回或设置图表的图表样式。可读/写 **Variant** 类型。

**语法**

**express.ChartStyle**

*express*   一个代表 **Chart** 对象的变量。

**说明**

您可以使用介于 1 到 48 之间的数字设置图表样式。

#### **Chart.ChartTitle**

返回一个 **ChartTitle** 对象，该对象表示指定图表的标题。只读。

**语法**

**express.ChartTitle**

*express*   一个代表 **Chart** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 Chart1 的标题文本。*/ function test() {     let charts = Application.Charts.Item("Chart1")     charts.HasTitle = true     charts.ChartTitle.Text = "First Quarter Sales" }` |

#### **Chart.ChartType**

返回或设置图表类型。**XlChartType** 类型，可读写。

**语法**

**express.ChartType**

*express*   一个代表 **Chart** 对象的变量。

**说明**

某些图表类型对数据透视图报表无效。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*当图表为二维气泡图时，本示例将第一个图表组中的气泡大小设置为默认大小的 200%。*/ function test() {     let chart = Application.Worksheets.Item(1).ChartObjects(1).Chart     if(chart.ChartType == xlBubble){         chart.ChartGroups(1).BubbleScale = 200     } }` |

#### **Chart.CodeName**

返回对象的代码名。**String** 型，只读。

**语法**

**express.CodeName**

*express*   一个代表 **Chart** 对象的变量。

**说明**

在“属性”窗口中“（名称）”右边的单元格中显示的值是所选对象的代码名。可以在设计过程中通过更改该值来改变对象的代码名。不能在运行过程中更改该属性。

对于一个返回指定对象的表达式，该表达式可使用对象的代码名。例如，如果第一张工作表的代码名为 Sheet1，则下列表达式是等价的。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Worksheets.Item(1).Range("a1") Sheet1.Range("a1")` |

工作表的名称可以与其代码名不同。创建一张工作表时，其工作表名称和代码名是相同的，不过，更改工作表名称时并不影响其代码名，并且，更改工作表代码名（在 Visual Basic 编辑器中使用“属性”窗口）也不影响其名称。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例显示第一张工作表的代码名。*/ alert(Application.Worksheets.Item(1).CodeName)` |

#### **Chart.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Chart** 对象的变量。

#### **Chart.DataTable**

返回一个 **DataTable** 对象，该对象表示图表的模拟运算表。只读。

**语法**

**express.DataTable**

*express*   一个代表 **Chart** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向嵌入图表添加带有外框的模拟运算表。*/ function test() {     let chart = Application.Worksheets.Item(1).ChartObjects(1).Chart     chart.HasDataTable = true     chart.DataTable.HasBorderOutline = true }` |

#### **Chart.DepthPercent**

返回或设置三维图表的深度，以图表宽度的百分比表示（有效范围从 20% 到 2000%）。**Long** 类型，可读写。

**语法**

**express.DepthPercent**

*express*   一个代表 **Chart** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 Chart1 的深度设置为其宽度的 50%。本示例应在三维图表上运行（DepthPercent 属性在二维图表中无效）。*/ Application.Charts.Item("Chart1").DepthPercent = 50` |

#### **Chart.DisplayBlanksAs**

返回或设置在图表中绘制空白单元格的方式。可以是 **XlDisplayBlanksAs** 常量之一。**Long** 类型，可读写。

**语法**

**express.DisplayBlanksAs**

*express*   一个代表 **Chart** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 ET 在 Chart1 中不绘制空白单元格。*/ Application.Charts.Item("Chart1").DisplayBlanksAs = xlNotPlotted` |

#### **Chart.Elevation**

返回或设置三维图表视图的仰角（以角度为单位）。**Long** 类型，可读写。

**语法**

**express.Elevation**

*express*   一个代表 **Chart** 对象的变量。

**说明**

图表的仰角以角度为单位表示观察图表的高度。对于绝大多数图表类型，其默认值为 15。本属性的值必须介于 -90 到 90 之间，但三维条形图除外，对于三维条形图，该值必须介于 0 到 44 之间。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 Chart1 的仰角设为 34 度。本示例应在三维图表上运行（Elevation 属性在二维图表上无效）。*/ Application.Charts.Item("Chart1").Elevation = 34` |

#### **Chart.Floor**

返回一个 **Floor** 对象，该对象表示三维图表的基底。只读。

返回值类型为Floor。

**语法**

**express.Floor**

*express*   一个代表 **Chart** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 Chart1 的基底颜色设为蓝色。本示例应在三维图表上运行（Floor 属性在二维图表无效）。*/ Application.Charts.Item("Chart1").Floor.Interior.ColorIndex = 5` |

#### **Chart.GapDepth**

以数据标志宽度的形式返回或设置三维图表中数据系列之间的距离，本属性的值必须在 0 和 500 之间。**Long** 类型，可读写。

**语法**

**express.GapDepth**

*express*   一个代表 **Chart** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 Chart1 的数据系列之间的距离设为数据标志宽度的 200%。本示例应在三维图表上运行（GapDepth 属性在二维图表上无效）。*/ Application.Charts.Item("Chart1").GapDepth = 2000` |

#### **Chart.HasAxis**

返回或设置图表上显示的坐标轴。**Variant** 类型，可读写。

**语法**

**express.HasAxis**

*express*   一个代表 **Chart** 对象的变量。

**说明**

设置该属性时，必须至少输入一个参数的值。

如果更改了图表类型或 **Axis.AxisGroup**、**Chart.AxisGroup** 或 **Series.AxisGroup** 属性，ET 可能要创建或删除坐标轴。

参数

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Index1* | 必选          | **Variant**  | 坐标轴类型。系列坐标轴仅适用于三维图表。可以是 **XlAxisType** 常量之一。 |
| *Index2* | 可选          | **Variant**  | 坐标轴组。三维图表只有一组坐标轴。可以是 **XlAxisGroup** 常量之一。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示 Chart1 中的主数值轴。*/ Application.Charts.Item("Chart1").HasAxis(xlValue, xlPrimary) = true` |

#### **Chart.HasDataTable**

如果图表有模拟运算表，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.HasDataTable**

*express*   一个代表 **Chart** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例使嵌入图表模拟运算表显示时带有外边框，但无单元格边框。*/ function test() {     let chart = Application.Worksheets.Item(1).ChartObjects(1).Chart     chart.HasDataTable = true     let newchart = chart.DataTable         newchart.HasBorderHorizontal = false         newchart.HasBorderVertical = false         newchart.HasBorderOutline = true }` |

#### **Chart.HasLegend**

如果图表有图例，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.HasLegend**

*express*   一个代表 **Chart** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示 Chart1 的图例，然后将图例的字体颜色设为蓝色。*/ function test() {     let charts = Application.Charts.Item("Chart1")     charts.HasLegend = true     charts.Legend.Font.ColorIndex = 5 }` |

#### **Chart.HasTitle**

如果坐标轴或图表有可见标题，则为 **True**。**Boolean** 类型，可读写。

**语法**

**express.HasTitle**

*express*   一个代表 **Chart** 对象的变量。

**说明**

图表标题由 **ChartTitle** 对象代表。

#### **Chart.HeightPercent**

返回或设置三维图表的高度，以图表宽度的百分比表示（有效范围从 5% 到 500%）。**Long** 类型，可读写。

**语法**

**express.HeightPercent**

*express*   一个代表 **Chart** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 Chart1 的高度设为其宽度的 80%。本示例应在三维图表上运行。*/ Application.Charts.Item("Chart1").HeightPercent = 80` |

#### **Chart.Hyperlinks**

返回一个 **Hyperlinks** 集合，它代表图表的超链接。

**语法**

**express.Hyperlinks**

*express*   一个代表 **Chart** 对象的变量。

#### **Chart.Index**

返回 **Long** 值，它代表对象在其同类对象所组成的集合内的索引号。

**语法**

**express.Index**

*express*   一个代表 **Chart** 对象的变量。

#### **Chart.Legend**

返回一个 **Legend** 对象，该对象表示指定图表的图例。只读。

**语法**

**express.Legend**

*express*   一个代表 **Chart** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示 Chart1 的图例，然后将图例的字体颜色设为蓝色。*/ function test() {     Application.Charts.Item("Chart1").HasLegend = true     Application.Charts.Item("Chart1").Legend.Font.ColorIndex = 5 }` |

#### **Chart.MailEnvelope**

代表文档的电子邮件标题。

**语法**

**express.MailEnvelope**

*express*   一个代表 **Chart** 对象的变量。

#### **Chart.Name**

返回或设置一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

*express*   一个代表 **Chart** 对象的变量。

**说明**

此属性对图表对象（嵌入式图表）而言是只读的。

#### **Chart.Next**

返回代表下一个工作表的 **Worksheet** 对象。

**语法**

**express.Next**

*express*   一个代表 **Chart** 对象的变量。

**说明**

如果该对象为区域，则此属性会模拟 Tab 键，但会返回下一个单元格，而不选中它。

在处于保护状态的工作表中，此属性返回下一个未锁定单元格。在未保护的工作表中，此属性总是返回紧靠指定单元格右边的单元格。

#### **Chart.PageSetup**

返回一个 **PageSetup** 对象，它包含用于指定对象的所有页面设置。只读。

**语法**

**express.PageSetup**

*express*   一个代表 **Chart** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例设置 Chart1 的中央页眉的文字。*/ Application.Charts.Item("Chart1").PageSetup.CenterHeader = "December Sales"` |

#### **Chart.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **Chart** 对象的变量。

#### **Chart.Perspective**

返回或设置一个 **Long** 值，它代表三维图表视图的透视系数。

**语法**

**express.Perspective**

*express*   一个代表 **Chart** 对象的变量。

**说明**

此属性的值必须在 0 到 100 之间。如果 **RightAngleAxes** 属性是 **True**，则此属性会被忽略。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例设置 Chart1 的透视系数为 70。此示例应在三维图表上运行。*/ function test() {     Application.Charts.Item("Chart1").RightAngleAxes = false     Application.Charts.Item("Chart1").Perspective = 70 }` |

#### **Chart.PivotLayout**

返回一个 **PivotLayout** 对象，该对象表示数据透视表中字段的位置以及数据透视图中坐标轴的位置。只读。

**语法**

**express.PivotLayout**

*express*   一个代表 **Chart** 对象的变量。

**说明**

如果指定图表不是一个数据透视图，则该属性值为 **Nothing**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例创建第一个数据透视图中使用的所有数据透视表字段的名称列表。*/ function test() {     let objNewSheet = Application.Worksheets.Add()     objNewSheet.Activate()     let intRow = 1     for(let i =1; i <= Application.Charts.Item("Chart1").PivotLayout.PivotFields().Count; i++){         objNewSheet.Cells.Item(intRow, 1).Value2 = Application.Charts.Item("Chart1").PivotLayout.PivotFields(i).Caption         intRow ++     } }` |

#### **Chart.PlotArea**

返回一个 **PlotArea** 对象，该对象表示图表的绘图区。只读。

**语法**

**express.PlotArea**

*express*   一个代表 **Chart** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 Chart1 绘图区的内部颜色设置为蓝绿色。*/ Application.Charts.Item("Chart1").PlotArea.Interior.ColorIndex = 8` |

#### **Chart.PlotBy**

返回或设置行或列在图表中作为数据系列使用的方式。可为以下 **XlRowCol** 常量之一：**xlColumns** 或 **xlRows**。**Long** 类型，可读写。

**语法**

**express.PlotBy**

*express*   一个代表 **Chart** 对象的变量。

**说明**

对于数据透视图，该属性是只读的，并且总返回 **xlColumns**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例使嵌入图表按列绘制数据。*/ Application.Worksheets.Item(1).ChartObjects(1).Chart.PlotBy = xlColumns` |

#### **Chart.PlotVisibleOnly**

如果仅绘制可见单元格，则该值为 **True**。如果可见单元格和隐藏单元格都绘制，则该值为 **False**。**Boolean** 类型，可读写。

**语法**

**express.PlotVisibleOnly**

*express*   一个代表 **Chart** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例使 ET 在 Chart1 中仅绘制可见单元格。*/ Application.Charts.Item("Chart1").PlotVisibleOnly = true` |

#### **Chart.Previous**

返回代表下一个工作表的 **Worksheet** 对象。

**语法**

**express.Previous**

*express*   一个代表 **Chart** 对象的变量。

**说明**

如果指定对象为区域，则此属性的作用是仿效 Shift+Tab；但此属性只是返回上一单元格，并不选定它。

在保护工作表上，该属性返回上一个未锁定的单元格。在未保护的工作表上，该属性通常返回指定单元格左侧相邻的单元格。

#### **Chart.PrintedCommentPages**

返回将为当前图表打印的批注页的数量。只读。

返回类型类型为**Long。**

**语法**

**express.PrintedCommentPages**

*express*   一个代表 **Chart** 对象的变量。

**说明**

由于图表和图表工作表不支持批注，因此 **Chart** 对象的 **PrintedCommentPages** 属性将总是返回零。

#### **Chart.ProtectContents**

如果工作表内容是受保护的，则为 **True**。对于图表，这样会保护整个图表。要打开内容保护，请使用 **Protect** 方法，并将 *Contents* 参数设置为 **True**。**Boolean** 类型，只读。

**语法**

**express.ProtectContents**

*express*   一个代表 **Chart** 对象的变量。

#### **Chart.ProtectData**

如果用户不能更改系列公式，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ProtectData**

*express*   一个代表 **Chart** 对象的变量。

**说明**

该属性在文件被保存后将改变。如果将该属性设置为 **True** 然后重新打开文件，则它不再为 **True**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例保护第一张工作表上嵌入的第一个图表中的数据。*/ Application.Worksheets.Item(1).ChartObjects(1).Chart.ProtectData = true` |

#### **Chart.ProtectDrawingObjects**

如果形状是受保护的，则为 **True**。要打开形状保护，请使用 **Protect** 方法，并将 *DrawingObjects* 参数设置为 **True**。**Boolean** 类型，只读。

**语法**

**express.ProtectDrawingObjects**

*express*   一个代表 **Chart** 对象的变量。

#### **Chart.ProtectFormatting**

如果用户不能更改格式，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ProtectFormatting**

*express*   一个代表 **Chart** 对象的变量。

**说明**

该属性在文件被保存后将改变。如果将该属性设置为 **True** 然后重新打开文件，则它不再为 **True**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例保护第一张工作表上嵌入的第一个图表的格式。*/ Application.Worksheets.Item(1).ChartObjects(1).Chart.ProtectFormatting = true` |

#### **Chart.ProtectSelection**

如果不能选定图表元素，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ProtectSelection**

*express*   一个代表 **Chart** 对象的变量。

**说明**

该属性为 **True** 时，不能向图表中添加形状，而且图表元素的 **Click** 和 **DoubleClick** 事件不会发生。

该属性在文件被保存后将改变。如果将该属性设置为 **True** 然后重新打开文件，则它不再为 **True**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例使第一张工作表上嵌入的第一个图表中的图表元素不能被选定。*/ Application.Worksheets.Item(1).ChartObjects(1).Chart.ProtectSelection = true` |

#### **Chart.ProtectionMode**

如果启用了用户界面专用保护，则为 **True**。要打开用户界面保护，请使用 **Protect** 方法，并将 *UserInterfaceOnly* 参数设置为 **True**。**Boolean** 类型，只读。

**语法**

**express.ProtectionMode**

*express*   一个代表 **Chart** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例显示 ProtectionMode 属性的状态。*/ alert(Application.ActiveSheet.ProtectionMode)` |

#### **Chart.RightAngleAxes**

如果图表的坐标轴为直角，并与图表的转角或仰角无关，则该值为 **True**。仅应用于三维折线图、柱形图和条形图。**Boolean** 类型，可读写。

**语法**

**express.RightAngleAxes**

*express*   一个代表 **Chart** 对象的变量。

**说明**

如果该属性值为 **True**，则忽略 **Perspective** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 Chart1 中坐标轴夹角为直角。本示例应在三维图表上运行。*/ Application.Charts.Item("Chart1").RightAngleAxes = true` |

#### **Chart.Rotation**

以度为单位返回或设置三维图表视图的转角（绘图区绕 Z 轴的转角）。此属性的取值必须介于 0 到 360 之间，三维条形图除外（从 0 到 44 之间）。默认值是 20。仅适用于三维图表。**Variant** 型，可读写。

**语法**

**express.Rotation**

*express*   一个代表 **Chart** 对象的变量。

**说明**

旋转量始终四舍五入到最接近的整数。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例将第一张图表的转角设置为 30 度。此示例应在三维图表上运行。*/ Application.Charts.Item("Chart1").Rotation = 30` |

#### **Chart.Shapes**

返回一个 **Shapes** 集合，它代表图表工作表上所有的形状。只读。

**语法**

**express.Shapes**

*express*   一个代表 **Chart** 对象的变量。

#### **Chart.ShowAllFieldButtons**

返回或设置是否在数据透视图上显示所有字段按钮。可读/写。

返回值类型为**Boolean。**

**语法**

**express.ShowAllFieldButtons**

*express*   一个代表 **Chart** 对象的变量。

**说明**

将 **ShowAllFieldButtons** 属性设置为 **True** 将在指定的数据透视图上显示所有字段按钮。将该属性设置为 **False** 将隐藏所有字段按钮。

**ShowAllFieldButtons** 属性对应于**“分析”**选项卡（在选择数据透视图时可用）的**“字段按钮”**下拉列表中的**“全部隐藏”**命令。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例将 Chart 1 设置为显示所有字段按钮。*/ function test() {     Application.ActiveSheet.ChartObjects("Chart 1").Activate()     Application.ActiveChart.ShowAllFieldButtons = true }` |

#### **Chart.ShowAxisFieldButtons**

返回或设置是否要在数据透视图上显示坐标轴字段按钮。可读/写。

返回值类型为**Boolean。**

**语法**

**express.ShowAxisFieldButtons**

*express*   一个代表 **Chart** 对象的变量。

**说明**

将 **ShowAxisFieldButtons** 属性设置为 **True** 将在指定的数据透视图上显示坐标轴字段按钮。将该属性设置为 **False** 将隐藏这些按钮。

**ShowAxisFieldButtons** 属性对应于**“分析”**选项卡（在选择数据透视图时可用）的**“字段按钮”**下拉列表中的**“显示坐标轴字段按钮”**命令。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例将 Chart 1 设置为隐藏坐标轴字段按钮。*/ function test() {     Application.ActiveSheet.ChartObjects("Chart 1").Activate()     Application.ActiveChart.ShowAxisFieldButtons = false }` |

#### **Chart.ShowDataLabelsOverMaximum**

返回或设置一个布尔值，该值表示在数值大于数值轴上的最大值时是否显示数据标签。可读/写 **Boolean** 类型。

**语法**

**express.ShowDataLabelsOverMaximum**

*express*   一个代表 **Chart** 对象的变量。

**说明**

如果您更改数值轴，使其变得比数据点的大小还小，则可以使用此属性设置是否显示数据标签。此属性只适用于二维图表。

#### **Chart.ShowLegendFieldButtons**

返回或设置是否要在数据透视图上显示图例字段按钮。可读/写。

返回值类型为**Boolean。**

**语法**

**express.ShowLegendFieldButtons**

*express*   一个代表 **Chart** 对象的变量。

**说明**

将 **ShowLegendFieldButtons** 属性设置为 **True** 将在指定的数据透视图上显示图例字段按钮。将该属性设置为 **False** 将隐藏这些按钮。

**ShowLegendFieldButtons** 属性对应于**“分析”**选项卡（在选择数据透视图时可用）的**“字段按钮”**下拉列表中的**“显示图例字段按钮”**命令。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例将 Chart 1 设置为显示图例字段按钮。*/ function test() {     Application.ActiveSheet.ChartObjects("Chart 1").Activate()     Application.ActiveChart.ShowLegendFieldButtons = true }` |

#### **Chart.ShowReportFilterFieldButtons**

返回或设置是否要在数据透视图上显示报表筛选字段按钮。可读/写。

返回值类型为**Boolean。**

**语法**

**express.ShowReportFilterFieldButtons**

*express*   一个代表 **Chart** 对象的变量。

**说明**

将 **ShowReportFilterFieldButtons** 属性设置为 **True** 将在指定的数据透视图上显示**“报表筛选字段”**按钮。将该属性设置为 **False** 将隐藏这些按钮。

**ShowReportFilterFieldButtons** 属性对应于**“分析”**选项卡（在选择数据透视图时可用）的**“字段按钮”**下拉列表中的**“显示报表筛选字段按钮”**命令。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例将 Chart 1 设置为显示报表筛选字段按钮。*/ function test() {     Application.ActiveSheet.ChartObjects("Chart 1").Activate()     Application.ActiveChart.ShowReportFilterFieldButtons = true }` |

#### **Chart.ShowValueFieldButtons**

返回或设置是否要在数据透视图上显示值字段按钮。可读/写。

返回值类型为**Boolean。**

**语法**

**express.ShowValueFieldButtons**

*express*   一个代表 **Chart** 对象的变量。

**说明**

将 **ShowValueFieldButtons** 属性设置为 **True** 将在指定的数据透视图上显示**“值字段”**按钮。将该属性设置为 **False** 将隐藏这些按钮。

**ShowValueFieldButtons** 属性对应于**“分析”**选项卡（在选择数据透视图时可用）的**“字段按钮”**下拉列表中的**“显示值字段按钮”**命令。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例将 Chart 1 设置为显示值字段按钮。*/ function test() {     Application.ActiveSheet.ChartObjects("Chart 1").Activate()     Application.ActiveChart.ShowValueFieldButtons = true }` |

#### **Chart.SideWall**

返回 **Walls** 对象，该对象允许用户单独对三维图表的侧面墙进行格式设置。只读。

**语法**

**express.SideWall**

*express*   一个代表 **Chart** 对象的变量。

#### **Chart.Tab**

为图表返回一个 **Tab** 对象。

**语法**

**express.Tab**

*express*   一个代表 **Chart** 对象的变量。

#### **Chart.Visible**

返回或设置一个 **XlSheetVisibility** 值，它确定对象是否可见。

**语法**

**express.Visible**

*express*   一个代表 **Chart** 对象的变量。

#### **Chart.Walls**

返回一个 [**Walls**](https://qn.cache.wpscdn.cn/encs/doc/office_v13/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Walls/Walls%20%E5%AF%B9%E8%B1%A1.htm) 对象，该对象表示三维图表的背景墙。只读。

**语法**

**express.Walls**

*express*   一个代表 **Chart** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 Chart1 的背景墙边框颜色设置为红色。本示例应在三维图表上运行。*/ Application.Charts.Item("Chart1").Walls.Border.ColorIndex = 3` |

适用环境：web

适用平台：windows/linux