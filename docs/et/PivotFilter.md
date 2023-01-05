**PivotFilter**



透视筛选应用于 **PivotField** 对象。

**说明**

由于索引不可靠，因此开发人员可以选择指定引用筛选。**DataField** 属性指定值筛选基于的透视字段。

**方法**

|                                                              | 名称       | 说明                                                   |
| ------------------------------------------------------------ | ---------- | ------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete** | 删除筛选并从透视字段和数据透视表的筛选集合中将其删除。 |

**属性**

|                                                              | 名称                       | 说明                                                         |
| ------------------------------------------------------------ | -------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Active**                 | table { word-break:break-all; }返回指定透视筛选是否是活动的。只读 **Boolean**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**            | table { word-break:break-all; }如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DataCubeField**          | table { word-break:break-all; }此属性仅适用于 OLAP 数据透视表，可为值筛选提供作为筛选依据的**值**字段（值区域中的透视字段）。可读写 **CubeField**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DataField**              | table { word-break:break-all; }此属性仅适用于非 OLAP 数据透视表，可为值筛选提供作为筛选依据的**值**字段（值区域中的透视字段）。可读写 **PivotField**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Description**            | table { word-break:break-all; }为 **PivotFilter** 对象提供可选说明。只读 **String**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FilterType**             | table { word-break:break-all; }指定要应用的筛选器类型。只读 **xlPivotFilterType** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IsMemberPropertyFilter** | table { word-break:break-all; }指定标签筛选器是基于字段的成员属性的 PivotItem 标题还是基于透视字段本身的 PivotItem 标题。只读 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MemberPropertyField**    | table { word-break:break-all; }此属性指定标签筛选器所基于的成员属性 PivotField。可读/写 **PivotField** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**                   | table { word-break:break-all; }此属性提供选项可指定引用筛选。索引值可能发生更改，因此不能依赖于这些值获得准确的引用。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Order**                  | table { word-break:break-all; }指定应用于整个数据透视表的所有“值”筛选器中筛选器的求值顺序。**Integer** 类型，可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                 | table { word-break:break-all; }返回指定 **PivotFilter** 对象的父对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PivotField**             | table { word-break:break-all; }指定应用筛选器的透视字段。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Value1**                 | table { word-break:break-all; }此属性是由用户提供的用于定义透视字段筛选器的参数。可读/写 **Variant** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Value2**                 | table { word-break:break-all; }此属性是由用户提供的用于定义透视字段筛选器的参数。可读/写 **Variant** 类型。 |

**成员方法**

#### **PivotFilter.Delete**

删除筛选并从透视字段和数据透视表的筛选集合中将其删除。

**语法**

**express.Delete()**

*express*   一个代表 **PivotFilter** 对象的变量。

**成员属性**

#### **PivotFilter.Active**

table { word-break:break-all; }

返回指定透视筛选是否是活动的。只读 **Boolean**。

**语法**

**express.Active**

*express*   一个代表 **PivotFilter** 对象的变量。

**说明**

table { word-break:break-all; }

当筛选的透视字段在数据透视表中，且数据透视表更新时会计算该筛选时，此属性返回 **True**。当筛选的透视字段不在数据透视表中，且对数据透视表计算没有影响时，此属性返回 **False**。

#### **PivotFilter.Application**

table { word-break:break-all; }

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **PivotFilter** 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

#### **PivotFilter.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **PivotFilter** 对象的变量。

**说明**

table { word-break:break-all; 

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **PivotFilter.DataCubeField**

table { word-break:break-all; }

此属性仅适用于 OLAP 数据透视表，可为值筛选提供作为筛选依据的**值**字段（值区域中的透视字段）。可读写 **CubeField**。

**语法**

**express.DataCubeField**

*express*   一个代表 **PivotFilter** 对象的变量。

#### **PivotFilter.DataField**

table { word-break:break-all; }

此属性仅适用于非 OLAP 数据透视表，可为值筛选提供作为筛选依据的**值**字段（值区域中的透视字段）。可读写 **PivotField**。

**语法**

**express.DataField**

*express*   一个代表 **PivotFilter** 对象的变量。

#### **PivotFilter.Description**

table { word-break:break-all; }

为 **PivotFilter** 对象提供可选说明。只读 **String**。

**语法**

**express.Description**

*express*   一个代表 **PivotFilter** 对象的变量。

**说明**

table { word-break:break-all; }

说明字符串的最大长度为 255 个字符。

#### **PivotFilter.FilterType**

table { word-break:break-all; }

指定要应用的筛选器类型。只读 **xlPivotFilterType** 类型。

**语法**

**express.FilterType**

*express*   一个代表 **PivotFilter** 对象的变量。

**说明**

table { word-break:break-all; }

下表列出了新的筛选器类型，以及对于每一个筛选器类型必须提供的参数和不可提供（不可用）的参数。

| xlPivotFilterType                 | DataField | DataCubeField (OLAP) | Value1 | Value2 |
| --------------------------------- | --------- | -------------------- | ------ | ------ |
| xlTopCount                        | 必需      | 必需                 | 必需   | 不可用 |
| xlBottomCount                     | 必需      | 必需                 | 必需   | 不可用 |
| xlTopPercent                      | 必需      | 必需                 | 必需   | 不可用 |
| xlBottomPercent                   | 必需      | 必需                 | 必需   | 不可用 |
| xlTopSum                          | 必需      | 必需                 | 必需   | 不可用 |
| xlBottomSum                       | 必需      | 必需                 | 必需   | 不可用 |
| xlCaptionEquals                   | 不可用    | 不可用               | 必需   | 不可用 |
| xlCaptionDoesNotEqual             | 不可用    | 不可用               | 必需   | 不可用 |
| xlCaptionIsGreaterThan            | 不可用    | 不可用               | 必需   | 不可用 |
| xlCaptionIsGreaterThanOrEqualTo   | 不可用    | 不可用               | 必需   | 不可用 |
| xlCaptionIsLessThan               | 不可用    | 不可用               | 必需   | 不可用 |
| xlCaptionIsLessThanOrEqualTo      | 不可用    | 不可用               | 必需   | 不可用 |
| xlCaptionBeginsWith               | 不可用    | 不可用               | 必需   | 不可用 |
| xlCaptionDoesNotBeginWith         | 不可用    | 不可用               | 必需   | 不可用 |
| xlCaptionEndsWith                 | 不可用    | 不可用               | 必需   | 不可用 |
| xlCaptionDoesNotEndWith           | 不可用    | 不可用               | 必需   | 不可用 |
| xlCaptionContains                 | 不可用    | 不可用               | 必需   | 不可用 |
| xlCaptionDoesNotContain           | 不可用    | 不可用               | 必需   | 不可用 |
| xlCaptionIsBetween                | 不可用    | 不可用               | 必需   | 必需   |
| xlCaptionIsNotBetween             | 不可用    | 不可用               | 必需   | 必需   |
| xlValueEquals                     | 必需      | 必需                 | 必需   | 不可用 |
| xlValueDoesNotEqual               | 必需      | 必需                 | 必需   | 不可用 |
| xlValueIsGreaterThan              | 必需      | 必需                 | 必需   | 不可用 |
| xlValueIsGreaterThanOrEqualTo     | 必需      | 必需                 | 必需   | 不可用 |
| xlValueIsLessThan                 | 必需      | 必需                 | 必需   | 不可用 |
| xlValueIsLessThanOrEqualTo        | 必需      | 必需                 | 必需   | 不可用 |
| xlValueIsBetween                  | 必需      | 必需                 | 必需   | 必需   |
| xlValueIsNotBetween               | 必需      | 必需                 | 必需   | 必需   |
| xlSpecificDate                    | 不可用    | 不可用               | 必需   | 不可用 |
| xlNotSpecificDate                 | 不可用    | 不可用               | 必需   | 不可用 |
| xlBefore                          | 不可用    | 不可用               | 必需   | 不可用 |
| xlBeforeOrEqualTo                 | 不可用    | 不可用               | 必需   | 不可用 |
| xlAfter                           | 不可用    | 不可用               | 必需   | 不可用 |
| xlAfterOrEqualTo                  | 不可用    | 不可用               | 必需   | 不可用 |
| xlBetween                         | 不可用    | 不可用               | 必需   | 不可用 |
| xlNotBetween                      | 不可用    | 不可用               | 必需   | 不可用 |
| xlFilterToday                     | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterYesterday                 | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterTomorrow                  | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterThisWeek                  | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterLastWeek                  | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterNextWeek                  | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterThisMonth                 | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterLastMonth                 | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterNextMonth                 | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterThisQuarter               | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterLastQuarter               | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterNextQuarter               | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterThisYear                  | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterLastYear                  | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterNextYear                  | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterYearToDate                | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodQuarter1  | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodQuarter2  | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodQuarter3  | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodQuarter4  | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodJanuary   | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodFebruary  | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodMarch     | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodApril     | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodMay       | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodJune      | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodJuly      | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodAugust    | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodSeptember | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodOctober   | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodNovember  | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodDecember  | 不可用    | 不可用               | 不可用 | 不可用 |

#### **PivotFilter.IsMemberPropertyFilter**

table { word-break:break-all; }

指定标签筛选器是基于字段的成员属性的 PivotItem 标题还是基于透视字段本身的 PivotItem 标题。只读 **Boolean** 类型。

**语法**

**express.IsMemberPropertyFilter**

*express*   一个代表 **PivotFilter** 对象的变量。

**说明**

table { word-break:break-all; }

此属性的默认值为 **False**。

如果标签筛选器基于透视字段的成员属性的 PivotItem 标题，则返回 **True**；或者，如果筛选器基于透视字段的 PivotItem 标题，则返回 **False**。此属性只对标签筛选器和至少有一个成员属性的 OLAP 透视字段有效。

#### **PivotFilter.MemberPropertyField**

table { word-break:break-all; }

此属性指定标签筛选器所基于的成员属性 PivotField。可读/写 **PivotField** 类型。

**语法**

**express.MemberPropertyField**

*express*   一个代表 **PivotFilter** 对象的变量。

**说明**

table { word-break:break-all; }

此属性只对标签筛选器和至少有一个成员属性的 OLAP 透视字段有效。

#### **PivotFilter.Name**

table { word-break:break-all; }

此属性提供选项可指定引用筛选。索引值可能发生更改，因此不能依赖于这些值获得准确的引用。

**语法**

**express.Name**

*express*   一个代表 **PivotFilter** 对象的变量。

#### **PivotFilter.Order**

table { word-break:break-all; }

指定应用于整个数据透视表的所有“值”筛选器中筛选器的求值顺序。**Integer** 类型，可读/写。

**语法**

**express.Order**

*express*   一个代表 **PivotFilter** 对象的变量。

**说明**

table { word-break:break-all; }

此属性仅对“值”和“前 *n* 个”类型的 PivotFilters 有效。如果尝试设置或获取“标签”和“日期”筛选器的该属性，则会返回一个运行时错误。1 代表已求值的第一个筛选器，2 代表已求值的下一个筛选器，依此类推，直至到达第 *n* 个值。-1 代表非活动筛选器。

如果在添加新的筛选器时未指定 **EvaluationOrder** 属性，则该属性将被设置为 *N+1*（其中 *N* 为筛选器集合中当前最大的 **EvaluationOrder** 数字）。

可在 **Add** 方法中指定此属性，或者可在稍后通过更改属性来设置字段的此属性。

如果提高某个字段的求值顺序，则会将原来具有该求值顺序值的字段（以及这两个字段之间的所有字段）的求值顺序降低 1。将求值顺序设置为和原来一样将不会产生变化。如果降低某个字段的求值顺序，则会将原来具有该求值顺序值的字段（以及这两个字段之间的所有字段）的求值顺序提高 1。

集合中 PivotFilters 的顺序与对其进行求值的顺序相同。因此开发人员可以更改对透视字段进行求值的顺序。如果将透视字段（非 OLAP 数据透视表）或 CubeField（OLAP 数据透视表）从 **PivotTables** 集合中删除，则会针对应用于字段的“值”或“前 *n* 个”筛选器，将此属性设置为 -1。如果未明确指定值，那么在将删除的字段添加回去时，会针对应用的“值”或“前 *n* 个”筛选器，将 **EvaluationOrder** 属性设置为 *N+1*。

#### **PivotFilter.Parent**

table { word-break:break-all; }

返回指定 **PivotFilter** 对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **PivotFilter** 对象的变量。

#### **PivotFilter.PivotField**

table { word-break:break-all; }

指定应用筛选器的透视字段。只读。

**语法**

**express.PivotField**

*express*   一个代表 **PivotFilter** 对象的变量。

#### **PivotFilter.Value1**

table { word-break:break-all; }

此属性是由用户提供的用于定义透视字段筛选器的参数。可读/写 **Variant** 类型。

**语法**

**express.Value1**

*express*   一个代表 **PivotFilter** 对象的变量。

#### **PivotFilter.Value2**

table { word-break:break-all; }

此属性是由用户提供的用于定义透视字段筛选器的参数。可读/写 **Variant** 类型。

**语法**

**express.Value2**

*express*   一个代表 **PivotFilter** 对象的变量。

适用环境：web

适用平台：windows/linux