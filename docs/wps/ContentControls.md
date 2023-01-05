#### **ContentControls**



**ContentControl**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/ContentControl/ContentControl%20.htm#jsObject_ContentControl)对象的集合。内容控件是文档中绑定的、有可能添加标签的区域，它们充当特定类型的内容的容器。单个内容控件可能包含诸如日期、列表或格式化文本段落等内容。

**方法**

|                                                              | 名称     | 说明                                                         |
| ------------------------------------------------------------ | -------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**  | 将指定类型的新内容控件添加到活动文档中，并返回表示新内容控件的 **ContentControl**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/ContentControl/ContentControl%20.htm#jsObject_ContentControl)对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 返回 **ContentControl** 对象，该对象表示文档中的内容控件集合内的指定内容控件。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个 **Application** 对象，该对象代表 WPS 应用程序。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回 **ContentControls** 集合中的项目数。只读 **Number** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位的整数，该整数表示在其中创建加载项的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回一个 **Object** 类型的值，该值代表指定 **ContentControls** 对象的父对象。 |

**成员方法**

#### **ContentControls.Add**

将指定类型的新内容控件添加到活动文档中，并返回表示新内容控件的 **ContentControl**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/ContentControl/ContentControl%20.htm#jsObject_ContentControl)对象。

**语法**

**express.Add(Type, Range)**

*express*   一个代表 **ContentControls** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**             | **说明**                                                     |
| -------- | ------------- | ------------------------ | ------------------------------------------------------------ |
| *Type*   | 可选          | **WdContentControlType** | 指定要插入活动文档中的内容控件的类型。如果省略它，WPS 将插入 RTF 内容控件。 |
| *Range*  | 可选          | **Variant**              | 指定内容控件在活动文档中的放置位置。如果省略它， WPS 会将内容控件放在插入点所在位置，或者替换当前所选内容。 |

**说明**

只能在 RTF 内容控件、生成块库内容控件和组内容控件中嵌套内容控件。如果插入点或当前所选项在不同类型的内容控件的内部，则此方法将引发错误。在这种情况下，可以移动插入点，也可以使用 *Range* 参数来指定文档内的位置。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 添加一个下拉列表控件 */ Application.ActiveDocument.ContentControls.Add(wdContentControlDropdownList)` |

#### **ContentControls.Item**

返回 **ContentControl** 对象，该对象表示文档中的内容控件集合内的指定内容控件。

**语法**

**express.Item(Index)**

*express*   一个代表 **ContentControls** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                         |
| -------- | ------------- | ------------ | -------------------------------- |
| *Index*  | 可选          | **Variant**  | 指定要返回的内容控件的序号位置。 |

**成员属性**

#### **ContentControls.Application**

返回一个 **Application** 对象，该对象代表 WPS 应用程序。

**语法**

**express.Application**

*express*   一个代表 **ContentControls** 对象的变量。

#### **ContentControls.Count**

返回 **ContentControls** 集合中的项目数。只读 **Number** 类型。

**语法**

**express.Count**

*express*   一个代表 **ContentControls** 对象的变量。

#### **ContentControls.Creator**

返回一个 32 位的整数，该整数表示在其中创建加载项的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **ContentControls** 对象的变量。

**说明**

如果对象是在 WPS 中创建的，则 **Creator** 属性返回十六进制数 4D535744，代表字符串“WPS”。该属性主要是为 Macintosh 机的应用设计的，在该机上每个应用程序都有一个四字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参考 WPS OfficeMacintosh Edition 中的语言参考帮助。

| 注释                                    |
| --------------------------------------- |
| 该值也可用常量 **wdCreatorCode** 表示。 |

#### **ContentControls.Parent**

返回一个 **Object** 类型的值，该值代表指定 **ContentControls** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **ContentControls** 对象的变量。

适用环境：web

适用平台：windows/linux