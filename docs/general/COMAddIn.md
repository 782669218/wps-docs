#### **COMAddIn**



代表 WPS Office宿主应用程序中的一个 COM 加载项。**COMAddIn** 对象是 **COMAddIns** 集合的成员。

**说明**

可以使用 COMAddIns.Item(index)，其中 index 既可以是一个能够返回 COMAddIns 集合中相应位置的加载项的序数值，也可以是代表指定 COM 加载项的 ProgID 的 String 值。

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 获取一个 **Application** 对象，代表 **COMAddIn** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Connect**     | 获取或设置指定的 **COMAddIn** 对象的连接状态。可读/写。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 获取一个 32 位整数，指示创建 **COMAddIn** 对象时所使用的应用程序。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Description** | 获取或设置指定的 **COMAddin** 对象的说明性 **String** 值。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Guid**        | 获取指定 **COMAddIn** 对象的类标识符 (CLSID)。只读。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Object**      | 获取或设置一个对象引用。可读/写。                            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 获取 **COMAddIn** 对象的 **Parent** 对象。只读。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ProgId**      | 获取指定 **COMAddIn** 对象的编程标识符 (ProgID)。只读。      |

**成员属性**

#### **COMAddIn.Application**

获取一个 **Application** 对象，代表 **COMAddIn** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

*express*   一个代表 **COMAddIn** 对象的变量。

#### **COMAddIn.Connect**

获取或设置指定的 **COMAddIn** 对象的连接状态。可读/写。

**语法**

**express.Connect**

*express*   一个代表 **COMAddIn** 对象的变量。

**说明**

如果该加载项是活动的，那么 **Connect** 属性返回 **True**，反之返回 **False**。活动加载项是经过注册并建立了连接的加载项；非活动加载项是经过注册但当前没有连接的加载项。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的示例显示一个消息框，并在其中指示第一个 COM 加载项是否已经注册以及当前是否已经连接。 function test() {     if(Application.COMAddIns.Item(1).Connect)      {         alert("The add-in is connected.")     }     else      {         alert("The add-in is not connected.")     } }` |

#### **COMAddIn.Creator**

获取一个 32 位整数，指示创建 **COMAddIn** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

*express*   一个代表 **COMAddIn** 对象的变量。

#### **COMAddIn.Description**

获取或设置指定的 **COMAddin** 对象的说明性 **String** 值。可读/写。

**语法**

**express.Description**

*express*   一个代表 **COMAddIn** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的示例显示用于绘图的 COM 加载项的说明文本。 alert("The description of this COMAddIn is \"" + Application.COMAddIns.Item("msodraa9.ShapeSelect").Description + "\"")` |

#### **COMAddIn.Guid**

获取指定 **COMAddIn** 对象的类标识符 (CLSID)。只读。

**语法**

**express.Guid**

*express*   一个代表 **COMAddIn** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的示例在消息框中显示 COMAddIns 集合中第一个 COM 加载项的 ProgID 和 CLSID。 alert("My ProgID is " + Application.COMAddIns.Item(1).ProgID + " and my CLSID is "+ Application.COMAddIns.Item(1).Guid) ` |

#### **COMAddIn.Object**

获取或设置一个对象引用。可读/写。

**语法**

**express.Object**

*express*   一个代表 **COMAddIn** 对象的变量。

**说明**

**Object** 属性是一个可读写属性，可存储任何对象引用。从这个角度来讲，它类似于某些 ActiveX 控件的通用 **Tag** 属性。

在某些情况下，**Object** 属性返回指定的 **COMAddIn** 对象所代表的对象。其他情况下，它默认返回 **Nothing**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的示例返回 COM 加载项 msodraa9.ShapeSelect 所代表的对象。 let objBaseObject = Application.COMAddIns.Item("msodraa9.ShapeSelect").Object` |

#### **COMAddIn.Parent**

获取 **COMAddIn** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **COMAddIn** 对象的变量。

#### **COMAddIn.ProgId**

获取指定 **COMAddIn** 对象的编程标识符 (ProgID)。只读。

**语法**

**express.ProgId**

*express*   一个代表 **COMAddIn** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的示例在消息框中显示 COMAddIns 集合中第一个 COM 加载项的 ProgID 和 CLSID。 alert("My ProgID is " + Application.COMAddIns.Item(1).ProgID + " and my CLSID is "+ Application.COMAddIns.Item(1).Guid)` |

适用环境：web

适用平台：windows/linux