**PivotFilters**



**PivotFilters** 对象是 **PivotFilter** 对象的集合。

**说明**

**PivotFilters** 集合包含用于添加新筛选器、计算集合中现有筛选器个数以及引用特定 **PivotFilter** 对象的属性和方法。

在以下示例中，将向位于当前活动单元格位置的 PivotField 添加新的 PivotFilter。

| 示例代码                                             |
| ---------------------------------------------------- |
| `ActiveCell.PivotField.PivotFilters.Add(xlThisWeek)` |

**方法**

|                                                              | 名称    | 说明                                   |
| ------------------------------------------------------------ | ------- | -------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add** | 将新筛选添加到 **PivotFilters** 集合。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | table { word-break:break-all; }如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | table { word-break:break-all; }返回 **PivotFilters** 集合中的项数。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | table { word-break:break-all; }返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Item**        | table { word-break:break-all; }根据 **PivotFilters** 集合对象特定元素在集合中的位置返回该特定元素。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | table { word-break:break-all; }返回指定 **PivotFilters** 对象的父对象。只读。 |

**成员方法**

#### **PivotFilters.Add**

将新筛选添加到 **PivotFilters** 集合。

**语法**

**express.Add(Type, DataField, Value1, Value2, Order, Name, Description, MemberPropertyField)**

*express*   一个代表 **PivotFilters** 对象的变量。

**参数**

| **名称**              | **必选/可选** | **数据类型**          | **说明**                            |
| --------------------- | ------------- | --------------------- | ----------------------------------- |
| *Type*                | 必选          | **XlPivotFilterType** | 要求 XlPivotFilterType 类型的筛选。 |
| *DataField*           | 可选          | **Variant**           | 筛选附加到的字段。                  |
| *Value1*              | 可选          | **Variant**           | 筛选值 1。                          |
| *Value2*              | 可选          | **Variant**           | 筛选值 1。                          |
| *Order*               | 可选          | **Variant**           | 筛选数据的顺序。                    |
| *Name*                | 可选          | **Variant**           | 筛选的名称。                        |
| *Description*         | 可选          | **Variant**           | 筛选的简要说明                      |
| *MemberPropertyField* | 可选          | **Variant**           | 指定标签筛选所基于的成员属性字段。  |

**返回值**

PivotFilter

**示例**

下面是一些关于如何正确使用 **Add** 函数的示例。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `ActiveCell.PivotField.PivotFilters.Add(xlThisWeek)  ActiveCell.PivotField.PivotFilters.Add(xlTopCount,MyPivotField2,10)  ActiveCell.PivotField.PivotFilters.Add(xlCaptionIsNotBetween,null,"A","G")  ActiveCell.PivotField.PivotFilters.Add(xlValueIsGreaterThanOrEqualTo,MyPivotField2,10000) ` |

由于 *Value1* 的数据类型无效，本示例将返回一个运行时错误。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `ActiveCell.PivotField.PivotFilters.Add(xlValueIsGreaterThanOrEqualTo,MyPivotField2,“Allan”)` |

**成员属性**

#### **PivotFilters.Application**

table { word-break:break-all; }

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **PivotFilters** 对象的变量。

**说明**

table { word-break:break-all; }

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

#### **PivotFilters.Count**

table { word-break:break-all; }

返回 **PivotFilters** 集合中的项数。只读。

**语法**

**express.Count**

*express*   一个代表 **PivotFilters** 对象的变量。

#### **PivotFilters.Creator**

table { word-break:break-all; }

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **PivotFilters** 对象的变量。

**说明**

table { word-break:break-all; }

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **PivotFilters.Item**

table { word-break:break-all; }

根据 **PivotFilters** 集合对象特定元素在集合中的位置返回该特定元素。只读。

**语法**

**express.Item**

*express*   一个代表 **PivotFilters** 对象的变量。

#### **PivotFilters.Parent**

table { word-break:break-all; }

返回指定 **PivotFilters** 对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **PivotFilters** 对象的变量。

适用环境：web

适用平台：windows/linux