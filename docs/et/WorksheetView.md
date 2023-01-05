**WorksheetView**



指定的或活动工作簿中所有 **WorksheetView** 对象的集合。

**说明**

通过使用 DisplayFormulas、DisplayGridlines 和 DisplayHeadings 等属性控制应用程序或工作簿级视图的外观和风格。

**属性**

|                                                              | 名称                 | 说明                                                         |
| ------------------------------------------------------------ | -------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**      | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 [Application ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayFormulas**  | 返回或设置在当前工作表视图中是显示还是隐藏公式。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayGridlines** | 如果显示网格线，则为 **True**。可读/写 **Boolean** 类型。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayHeadings**  | 如果同时显示行标题和列标题，则为 **True**；如果未显示标题，则为 **False**。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayOutline**   | 如果显示分级显示符号，则为 **True**。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayZeros**     | 如果显示零值，则为 **True**。可读/写 **Boolean** 类型。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Sheet**            | 返回指定**WorksheetView**对象的工作表名称。只读。            |

**成员属性**

#### **WorksheetView.Application**

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 [Application ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **WorksheetView** 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

#### **WorksheetView.DisplayFormulas**

返回或设置在当前工作表视图中是显示还是隐藏公式。可读/写 **Boolean** 类型。

**语法**

**express.DisplayFormulas**

*express*   一个代表 **WorksheetView** 对象的变量。

#### **WorksheetView.DisplayGridlines**

如果显示网格线，则为 **True**。可读/写 **Boolean** 类型。

**语法**

**express.DisplayGridlines**

*express*   一个代表 **WorksheetView** 对象的变量。

**说明**

该属性仅适用于工作表和宏工作表。

此属性仅影响显示的网格线。使用**PrintGridlines**属性可以控制网格线的打印。

#### **WorksheetView.DisplayHeadings**

如果同时显示行标题和列标题，则为 **True**；如果未显示标题，则为 **False**。可读/写 **Boolean** 类型。

**语法**

**express.DisplayHeadings**

*express*   一个代表 **WorksheetView** 对象的变量。

**说明**

该属性仅适用于工作表和宏工作表。

此属性仅影响显示的标题。使用**PrintHeadings**属性可以控制标题的打印。

#### **WorksheetView.DisplayOutline**

如果显示分级显示符号，则为 **True**。可读/写 **Boolean** 类型。

**语法**

**express.DisplayOutline**

*express*   一个代表 **WorksheetView** 对象的变量。

**说明**

该属性仅适用于工作表和宏工作表。

#### **WorksheetView.DisplayZeros**

如果显示零值，则为 **True**。可读/写 **Boolean** 类型。

**语法**

**express.DisplayZeros**

*express*   一个代表 **WorksheetView** 对象的变量。

**说明**

该属性仅适用于工作表和宏工作表。

#### **WorksheetView.Sheet**

返回指定**WorksheetView**对象的工作表名称。只读。

**语法**

**express.Sheet**

*express*   一个代表 **WorksheetView** 对象的变量。

适用环境：web

适用平台：windows/linux