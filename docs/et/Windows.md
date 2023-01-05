**Windows**



ET 中所有 **Window** 对象的集合。

**说明**

**Application** 对象的 **Windows** 集合包含应用程序中的所有窗口，而 **Workbook** 对象的 **Windows** 集合只包含指定工作簿中的窗口。

## 示例

使用 **Windows** 属性可返回 **Windows** 集合。下例层叠当前在 ET 中显示的所有窗口。

| 示例代码复制                             |
| ---------------------------------------- |
| `Application.Windows.Arrange(xlCascade)` |

使用 **NewWindow** 方法可返回一个新窗口并将它添加到集合。下例为活动工作簿创建一个新窗口。

| 示例代码复制                             |
| ---------------------------------------- |
| `Application.ActiveWorkbook.NewWindow()` |

使用 **Windows**(*index*)（其中 *index* 是窗口名称或索引号）可返回一个 **Window** 对象。下例最大化活动窗口。

注意，活动窗口总是 `Windows(1)`。

| 示例代码复制                                            |
| ------------------------------------------------------- |
| `Application.Windows.Item(1).WindowState = xlMaximized` |

**方法**

|                                                              | 名称     | 说明                   |
| ------------------------------------------------------------ | -------- | ---------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 从集合中返回一个对象。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Long** 值，它代表集合中对象的数量。               |

**成员方法**

#### **Windows.Item**

从集合中返回一个对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **Windows** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**             |
| -------- | ------------- | ------------ | -------------------- |
| *Index*  | 必选          | **Variant**  | 对象的名称或索引号。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例将当前窗口最大化。*/ Application.Windows.Item(1).WindowState = xlMaximized` |

**成员属性**

#### **Windows.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **Windows** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() {     if (Application.ActiveWorkbook.Application.Value == "ET" ) {         alert("This is an ET Application object.")     }     else {         alert("This is not an ET Application object.")     } }` |

#### **Windows.Count**

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

*express*   一个代表 **Windows** 对象的变量。

适用环境：web

适用平台：windows/linux