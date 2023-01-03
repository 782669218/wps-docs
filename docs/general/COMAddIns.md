#### **COMAddIns**



**COMAddIn** 对象的集合，这些对象提供有关在 Windows 注册表中注册的 COM 加载项的信息。

**说明**

使用 **Application** 对象的 **COMAddIns** 属性可返回 WPS Office宿主应用程序的 **COMAddIns** 集合。如下面的示例所示，此集合包含了某个给定的 Office 宿主应用程序中所有可用的 COM 加载项，而 **COMAddins** 集合的 **Count** 属性可返回可用的 COM 加载项的数目。

| 示例代码                             |
| ------------------------------------ |
| `alert(Application.COMAddIns.Count)` |

如下面的示例所示，使用 **COMAddins** 集合的 **Update** 方法可刷新 Windows 注册表中 COM 加载项的列表。

| 示例代码                         |
| -------------------------------- |
| `Application.COMAddIns.Update()` |

可以使用 **COMAddIns.Item(index)**，其中 *index* 既可以是一个能够返回 **COMAddIns** 集合中相应位置的加载项的序数值，也可以是代表指定 COM 加载项的 ProgID 的 **String** 值。下面的示例在消息框中显示一个 COM 加载项的说明文本和 ProgID ("**msodraa9.ShapeSelect**")。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `alert(Application.COMAddIns.Item("msodraa9.ShapeSelect").Description)` |

**方法**

|                                                              | 名称       | 说明                                                         |
| ------------------------------------------------------------ | ---------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**   | 获取指定的 **COMAddIns** 集合的成员。                        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Update** | 根据保存在 Windows 注册表中的加载项列表更新 COMAddIns 集合的内容。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 获取一个 **Application** 对象，代表 **COMAddIns** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 获取宿主应用程序中 COM 加载项的数量。只读。                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 获取一个 32 位整数，指示创建 **COMAddIns** 对象时所使用的应用程序。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 获取 **COMAddIns** 对象的 **Parent** 对象。只读。            |

**成员方法**

#### **COMAddIns.Item**

获取指定的 **COMAddIns** 集合的成员。

**语法**

**express.Item(Index)**

*express*   一个代表 **COMAddIns** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**               |
| -------- | ------------- | ------------ | ---------------------- |
| *Index*  | 必选          | **Long**     | 代表集合中成员的位置。 |

#### **COMAddIns.Update**

根据保存在 Windows 注册表中的加载项列表更新 COMAddIns 集合的内容。

**语法**

**express.Update()**

*express*   一个代表 **COMAddIns** 对象的变量。

**说明**

在 WPS Office应用程序中使用特定 COM 加载项之前，必须先用相应的组件类别 ID 在 Windows 注册表中将该加载项注册为 COM 组件。通常，COM 加载项的安装程序会向注册表中添加必要的项。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的示例将根据保存在 Windows 注册表中的加载项列表更新 COMAddIns 集合的内容。 Application.COMAddIns.Update()` |

**成员属性**

#### **COMAddIns.Application**

获取一个 **Application** 对象，代表 **COMAddIns** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

*express*   一个代表 **COMAddIns** 对象的变量。

#### **COMAddIns.Count**

获取宿主应用程序中 COM 加载项的数量。只读。

**语法**

**express.Count**

*express*   一个代表 **COMAddIns** 对象的变量。

#### **COMAddIns.Creator**

获取一个 32 位整数，指示创建 **COMAddIns** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

*express*   一个代表 **COMAddIns** 对象的变量。

#### **COMAddIns.Parent**

获取 **COMAddIns** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **COMAddIns** 对象的变量。

适用环境：web

适用平台：windows/linux