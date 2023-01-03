**ChartView**



代表图表的视图。

**说明**

**ChartView** 对象是可由 **SheetViews** 集合（类似于 **Sheets** 集合）返回的对象之一。**ChartView** 对象只适用于图表工作表。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例返回 ChartView 对象。*/  ActiveWindow.SheetViews.Item(1) ` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例返回 Chart 对象。*/ ActiveWindow.SheetViews.Item(1).Sheet ` |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Sheet**       | 返回指定的 **ChartView** 对象的工作表名称。只读。            |

**成员属性**

#### **ChartView.Application**

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **ChartView** 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

#### **ChartView.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **ChartView** 对象的变量。

#### **ChartView.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **ChartView** 对象的变量。

#### **ChartView.Sheet**

返回指定的 **ChartView** 对象的工作表名称。只读。

**语法**

**express.Sheet**

*express*   一个代表 **ChartView** 对象的变量。

适用环境：web

适用平台：windows/linux