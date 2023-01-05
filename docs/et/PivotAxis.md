**PivotAxis**



**PivotAxis** 对象用于在数据透视表中进行不对称深化。

**说明**

**PivotAxis** 对象包含诸如 **PivotRowAxis** 和 **PivotRowAxis** 之类的属性，用于处理数据透视表中的行和列。

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定 **PivotAxis** 对象的父对象。只读。                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PivotLines**  | 返回附加到指定 **PivotAxis** 对象的 PivotLines。只读。       |

**成员属性**

#### **PivotAxis.Application**

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **PivotAxis** 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

#### **PivotAxis.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **PivotAxis** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **PivotAxis.Parent**

返回指定 **PivotAxis** 对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **PivotAxis** 对象的变量。

#### **PivotAxis.PivotLines**

返回附加到指定 **PivotAxis** 对象的 PivotLines。只读。

**语法**

**express.PivotLines**

*express*   一个代表 **PivotAxis** 对象的变量。

适用环境：web

适用平台：windows/linux