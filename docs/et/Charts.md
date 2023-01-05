**Charts**



指定的或活动工作簿中所有图表工作表的集合。

**方法**

|                                                              | 名称             | 说明                                    |
| ------------------------------------------------------------ | ---------------- | --------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**          | 新建图表工作表，并返回 **Chart** 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Copy**         | 将工作表复制到工作簿的另一位置。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**       | 删除对象。                              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**         | 从集合中返回一个对象。                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Move**         | 将图表移到工作簿的另一位置。            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PrintOut**     | 打印对象。                              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PrintPreview** | 按对象打印后的外观效果显示对象的预览。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Select**       | 选择对象。                              |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Long** 值，它代表集合中对象的数量。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HPageBreaks** | 返回一个 **HPageBreaks** 集合，它代表图表上的水平分页符。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **VPageBreaks** | 返回一个 **VPageBreaks** 集合，它代表工作表上的垂直分页符。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Visible**     | 返回或设置一个 **Variant** 值，它确定对象是否可见。          |

**成员方法**

#### **Charts.Add**

新建图表工作表，并返回 **Chart** 对象。

**语法**

**express.Add(Before, After, Count, NewLayout)**

*express*   一个代表 **Charts** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Before*    | 可选          | **Variant**  | 指定工作表的对象，新建的工作表将置于此工作表之前。           |
| *After*     | 可选          | **Variant**  | 指定工作表的对象，新建的工作表将置于此工作表之后。           |
| *Count*     | 可选          | **Variant**  | 要添加的工作表数。默认值为 1。                               |
| *NewLayout* | 可选          | **Variant**  | 如果 NewLayout 为 True， 则通过使用新的动态格式规则插入图表 (Title 为打开，并且 Legend 仅在有多个数据系列时) 。 |

**说明**

如果 *Before* 和 *After* 两者都被省略，新建的图表工作表将插入到活动工作表之前。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例创建空白图表工作表，并将其插入到最后一张工作表之前。*/ Application.ActiveWorkbook.Charts.Add(Worksheets.Item(Worksheets.Count))` |

#### **Charts.Copy**

将工作表复制到工作簿的另一位置。

**语法**

**express.Copy(Before, After)**

*express*   一个代表 **Charts** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Before* | 可选          | **Variant**  | 将要在其之前放置所复制工作表的工作表。如果指定了 After，则不能指定 Before。 |
| *After*  | 可选          | **Variant**  | 将要在其之后放置所复制工作表的工作表。如果指定了 Before，则不能指定 After。 |

**说明**

如果既不指定 *Before* 也不指定 *After*，则 ET 将新建一个工作簿，其中包含复制的工作表。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例复制工作表 Sheet1，并将其放置在工作表 Sheet3 之后。*/ Application.Worksheets.Item("Sheet1").Copy(null, Worksheets.Item("Sheet3"))` |

#### **Charts.Delete**

删除对象。

**语法**

**express.Delete()**

*express*   一个代表 **Charts** 对象的变量。

#### **Charts.Item**

从集合中返回一个对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **Charts** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**             |
| -------- | ------------- | ------------ | -------------------- |
| *Index*  | 必选          | **Variant**  | 对象的名称或索引号。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例对 Chart1 中的趋势线向前和向后延伸的单位数进行设置。此示例应在包含单个带趋势线系列的二维柱形图上运行。 */ function test() {     let myChartLine = Application.Charts.Item("Chart1").SeriesCollection(1).Trendlines(1)     myChartLine.Forward = 5     myChartLine.Backward = .5 } ` |

#### **Charts.Move**

将图表移到工作簿的另一位置。

**语法**

**express.Move(Before, After)**

*express*   一个代表 **Charts** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Before* | 可选          | **Variant**  | 将要在其之前放置所移动图表的工作表。如果指定了 After，则不能指定 Before。 |
| *After*  | 可选          | **Variant**  | 将要在其之后放置所移动图表的工作表。如果指定了 Before，则不能指定 After。 |

**说明**

如果既不指定 *Before* 也不指定 *After*，则 ET 将新建一个工作簿并将要移动的图表移到新工作簿中。

#### **Charts.PrintOut**

打印对象。

**语法**

**express.PrintOut(From, To, Copies, Preview, ActivePrinter, PrintToFile, Collate, PrToFileName, IgnorePrintAreas)**

*express*   一个代表 **Charts** 对象的变量。

**参数**

| **名称**           | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------------ | ------------- | ------------ | ------------------------------------------------------------ |
| *From*             | 可选          | **Variant**  | 打印的开始页号。如果省略此参数，则从起始位置开始打印。       |
| *To*               | 可选          | **Variant**  | 打印的终止页号。如果省略此参数，则打印至最后一页。           |
| *Copies*           | 可选          | **Variant**  | 打印份数。如果省略此参数，则只打印一份。                     |
| *Preview*          | 可选          | **Variant**  | 如果为 True，ET 将在打印对象之前调用打印预览。如果为 False（或省略该参数），则立即打印对象。 |
| *ActivePrinter*    | 可选          | **Variant**  | 设置活动打印机的名称。                                       |
| *PrintToFile*      | 可选          | **Variant**  | 如果为 True，则打印到文件。如果没有指定 PrToFileName，ET 将提示用户输入要使用的输出文件的文件名。 |
| *Collate*          | 可选          | **Variant**  | 如果为 True，则逐份打印多个副本。                            |
| *PrToFileName*     | 可选          | **Variant**  | 如果 PrintToFile 设为 True，则该参数指定要打印到的文件名。   |
| *IgnorePrintAreas* | 可选          | **Variant**  | 如果为 True，则忽略打印区域并打印整个对象。                  |

**说明**

*From* 和 *To* 所描述的“页”指的是要打印的页，并非指定工作表或工作簿中的全部页。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例打印当前活动的Charts工作表。*/ Application.Charts.PrintOut()` |

#### **Charts.PrintPreview**

按对象打印后的外观效果显示对象的预览。

**语法**

**express.PrintPreview(EnableChanges)**

*express*   一个代表 **Charts** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型** | **说明**                                                     |
| --------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *EnableChanges* | 可选          | **Variant**  | 传递 Boolean 值，以指定用户是否可更改边距和打印预览中可用的其他页面设置选项。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例在打印预览中显示Charts工作表。*/ Application.Charts.PrintPreview()` |

#### **Charts.Select**

选择对象。

**语法**

**express.Select(Replace)**

*express*   一个代表 **Charts** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型** | **说明**                                                     |
| --------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Replace* | 可选          | **Variant**  | （仅用于工作表）。如果为 True，则用指定的对象替换当前所选内容。如果为 False，则扩展当前所选内容以包括以前选择的对象和指定的对象。 |

**成员属性**

#### **Charts.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **Charts** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() {     let myObject = Application.ActiveWorkbook     if(myObject.Application.Value == "ET") {         alert("This is an ET Application object.")     }     else {         alert("This is not an ET Application object.")     } }` |

#### **Charts.Count**

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

*express*   一个代表 **Charts** 对象的变量。

#### **Charts.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Charts** 对象的变量。

#### **Charts.HPageBreaks**

返回一个 **HPageBreaks** 集合，它代表图表上的水平分页符。只读。

**语法**

**express.HPageBreaks**

*express*   一个代表 **Charts** 对象的变量。

#### **Charts.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **Charts** 对象的变量。

#### **Charts.VPageBreaks**

返回一个 **VPageBreaks** 集合，它代表工作表上的垂直分页符。只读。

**语法**

**express.VPageBreaks**

*express*   一个代表 **Charts** 对象的变量。

#### **Charts.Visible**

返回或设置一个 **Variant** 值，它确定对象是否可见。

**语法**

**express.Visible**

*express*   一个代表 **Charts** 对象的变量。

适用环境：web

适用平台：windows/linux