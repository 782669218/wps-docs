#### **Revisions**



Revision 对象 [的集合](https://docs.microsoft.com/zh-cn/office/vba/api/word.revision) ，这些对象代表区域或文档中标有修订标记的更改。

**方法**

|                                                              | 名称          | 说明                                                         |
| ------------------------------------------------------------ | ------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AcceptAll** | 接受文档或区域中的所有修订、删除所有修订标记并将更改合并到文档中。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**      | 返回集合 中的单个 Revisions 对象。                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **RejectAll** | 拒绝接受文档的修订集合                                       |

**属性**

|                                                              | 名称      | 说明                                                 |
| ------------------------------------------------------------ | --------- | ---------------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count** | 返回一 个 Long ，代表集合中的修订数。 此为只读属性。 |

**成员方法**

#### **Revisions.AcceptAll**

接受文档或区域中的所有修订、删除所有修订标记并将更改合并到文档中。

**语法**

**express.AcceptAll()**

*express*   一个代表 **Revisions** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//接受选中区域的所有修订的 Application.Selection.Range.Revisions.AcceptAll()` |

#### **Revisions.Item**

返回集合 中的单个 Revisions 对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **Revisions** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**   |
| -------- | ------------- | ------------ | ---------- |
| *Index*  | 可选          | **Long**     | 对象的索引 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//获取选区的第一个修订对象 Application.Selection.Reivesion.Item(1)` |

#### **Revisions.RejectAll**

拒绝接受文档的修订集合

**语法**

**express.RejectAll()**

*express*   一个代表 **Revisions** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//拒绝接受选区的修订集合 Application.Selection.Range.Revisions.RejectAll()` |

**成员属性**

#### **Revisions.Count**

返回一 个 Long ，代表集合中的修订数。 此为只读属性。

**语法**

**express.Count**

*express*   一个代表 **Revisions** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//获取激活文档的修订个数。 Application.ActiveDocument.Revisions.Count` |

适用环境：web

适用平台：windows/linux