**PivotLine**



table { word-break:break-all; }

**PivotLine** 对象是 ET 数据透视表中的行或列的线条。

**说明**

table { word-break:break-all; }

PivotLine 只包含可见项，因此 **PivotLine** 集合中不存在折叠的项目子项以及隐藏级别中的项目。

PivotLine 在所有位置始终具有一个 PivotItem。这意味着与普通 PivotLine 相比，代表数据透视表中分类汇总的 PivotLine 包含较少的 PivotItem。

**属性**

|                                                              | 名称               | 说明                                                         |
| ------------------------------------------------------------ | ------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**    | table { word-break:break-all; }如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**        | table { word-break:break-all; }返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LineType**       | table { word-break:break-all; }返回一个表示 PivotLine 类型的 **XlPivotLineType** 常量。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**         | table { word-break:break-all; }返回指定 **PivotLine** 对象的父对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PivotLineCells** | table { word-break:break-all; }返回 PivotLine 中 **PivotCell** 对象的集合。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Position**       | table { word-break:break-all; }返回或设置 **PivotLine** 对象的位置。只读。 |

**成员属性**

#### **PivotLine.Application**

table { word-break:break-all; }

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **PivotLine** 对象的变量。

**说明**

table { word-break:break-all; }

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

#### **PivotLine.Creator**

table { word-break:break-all; }

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **PivotLine** 对象的变量。

**说明**

table { word-break:break-all; }

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **PivotLine.LineType**

table { word-break:break-all; }

返回一个表示 PivotLine 类型的 **XlPivotLineType** 常量。只读。

**语法**

**express.LineType**

*express*   一个代表 **PivotLine** 对象的变量。

#### **PivotLine.Parent**

table { word-break:break-all; }

返回指定 **PivotLine** 对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **PivotLine** 对象的变量。

#### **PivotLine.PivotLineCells**

table { word-break:break-all; }

返回 PivotLine 中 **PivotCell** 对象的集合。只读。

**语法**

**express.PivotLineCells**

*express*   一个代表 **PivotLine** 对象的变量。

#### **PivotLine.Position**

table { word-break:break-all; }

返回或设置 **PivotLine** 对象的位置。只读。

**语法**

**express.Position**

*express*   一个代表 **PivotLine** 对象的变量。

适用环境：web

适用平台：windows/linux