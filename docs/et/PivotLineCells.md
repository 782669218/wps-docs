**PivotLineCells**



table { word-break:break-all; }

特定 PivotLine 的 **PivotCell** 对象的集合。

**说明**

table { word-break:break-all; }

使用 **PivotLineCells**(*index*) 方法可以返回或指定集合中特定 **PivotCell** 对象的位置。您也可以指定 **PivotField** 对象或 PivotField 的名称以返回单个 **PivotCell** 对象。

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | table { word-break:break-all; }如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | table { word-break:break-all; }返回 **PivotLineCells** 集合中的项数。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | table { word-break:break-all; }返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Item**        | table { word-break:break-all; }                              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | table { word-break:break-all; }返回指定 **PivotLineCells** 对象的父对象。只读。 |

**成员属性**

#### **PivotLineCells.Application**

table { word-break:break-all; }

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **PivotLineCells** 对象的变量。

**说明**

table { word-break:break-all; }

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

#### **PivotLineCells.Count**

table { word-break:break-all; }

返回 **PivotLineCells** 集合中的项数。只读。

**语法**

**express.Count**

*express*   一个代表 **PivotLineCells** 对象的变量。

#### **PivotLineCells.Creator**

table { word-break:break-all; }

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **PivotLineCells** 对象的变量。

**说明**

table { word-break:break-all; }

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **PivotLineCells.Item**

table { word-break:break-all; }

**语法**

**express.Item**

*express*   一个代表 **PivotLineCells** 对象的变量。

#### **PivotLineCells.Parent**

table { word-break:break-all; }

返回指定 **PivotLineCells** 对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **PivotLineCells** 对象的变量。

适用环境：web

适用平台：windows/linux