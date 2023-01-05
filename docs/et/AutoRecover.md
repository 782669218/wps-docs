**AutoRecover**



代表工作簿的自动恢复功能。

**说明**

**AutoRecover** 对象的属性确定备份所有文件的路径和时间间隔。

使用 **Application** 对象的 **AutoRecover** 属性可返回 **AutoRecover** 对象。

使用 **AutoRecover** 对象的 **Path** 属性可设置“自动恢复”文件的保存路径。

示例

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将“自动恢复”文件的路径设置为驱动器 C。*/ function test(){     Application.AutoRecover.Path = "C:\\" }` |

使用 **AutoRecover** 

对象的 **Time** 

属性可设置备份所有文件的时间间隔。

| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/close.gif)注释 |
| ------------------------------------------------------------ |
| **Time** 属性的单位是分钟。                                  |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function SetTime(){     Application.AutoRecover.Time = 5 }` |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Enabled**     | 如果启用对象，则为 **True**。**Boolean** 类型，可读写。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Path**        | 返回或设置一个 **String** 值，它代表 ET 将要存储“自动恢复”临时文件的位置的完整路径。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Time**        | 设置或返回 **AutoRecover** 对象的时间间隔。允许的值为从 1 到 120 分钟的整数。默认值为 10 分钟。**Long** 类型，可读写。 |

**成员属性**

#### **AutoRecover.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **AutoRecover** 对象的变量。

#### **AutoRecover.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **AutoRecover** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **AutoRecover.Enabled**

如果启用对象，则为 **True**。**Boolean** 类型，可读写。

**语法**

**express.Enabled**

*express*   一个代表 **AutoRecover** 对象的变量。

#### **AutoRecover.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **AutoRecover** 对象的变量。

#### **AutoRecover.Path**

返回或设置一个 **String** 值，它代表 ET 将要存储“自动恢复”临时文件的位置的完整路径。

**语法**

**express.Path**

*express*   一个代表 **AutoRecover** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例将“自动恢复”文件的路径设置为驱动器 C。*/ function test(){     Application.AutoRecover.Path = "C:\\" }` |

#### **AutoRecover.Time**

设置或返回 **AutoRecover** 对象的时间间隔。允许的值为从 1 到 120 分钟的整数。默认值为 10 分钟。**Long** 类型，可读写。

**语法**

**express.Time**

*express*   一个代表 **AutoRecover** 对象的变量。

**说明**

输入的小数值将舍入为最接近的整数。例如，输入的 5.5 与 6 等价。

如果输入有效区域以外的时间值，则 ET 将使用上一次使用的时间值。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例将“自动恢复”的时间间隔设置为 5 分钟，并通知用户。*/ function test(){     Application.AutoRecover.Time = 5     alert("The AutoRecover time interval is set at " + Application.AutoRecover.Time + " minutes.") }` |

适用环境：web

适用平台：windows/linux