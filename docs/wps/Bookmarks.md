**Bookmark** 

**Bookmark** 对象的集合，该集合代表指定选定内容、区域或文档中的书签。

**方法**

|                                                              | 名称       | 说明                                                        |
| ------------------------------------------------------------ | ---------- | ----------------------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**    | 返回一个 **Bookmark** 对象，该对象代表添加到区域中的书签。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Exists** | 判断指定书签是否存在。如果存在指定的书签，则返回 **True**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**   | 返回集合中的单个 **Bookmark** 对象。                        |

**属性**

|                                                              | 名称               | 说明                                                         |
| ------------------------------------------------------------ | ------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**    | 返回一个 [Application ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象代表 WPS 应用程序。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**          | 返回 **Bookmarks** 集合中的项目数。**Number** 类型，只读。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**        | 返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。**Long** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DefaultSorting** | 返回或设置**“插入”**菜单上**“书签”**对话框中显示的书签名的排序依据选项。**WdBookmarkSortBy** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**         | 返回一个 **Object** 类型的值，该值代表指定 **Bookmarks** 集合中的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowHidden**     | 如果 **Bookmarks** 集合中包括隐藏书签，则该属性值为 **True**。**Boolean** 类型，可读写。 |

**成员方法**

#### **Bookmarks.Add**

返回一个 **Bookmark** 对象，该对象代表添加到区域中的书签。

**语法**

**express.Add(Name, Range)**

*express*   一个代表 **Bookmarks** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Name*   | 必选          | **String**   | 书签名。书签名不能多于一个单词。                             |
| *Range*  | 可选          | **Variant**  | 书签标记的文本区域。可将书签设置到一个折叠的区域（插入点）。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 在当前插入点插入名为 "point" 的书签 */ Application.ActiveDocument.Bookmarks.Add("point",null)` |

#### **Bookmarks.Exists**

判断指定书签是否存在。如果存在指定的书签，则返回 **True**。

**语法**

**express.Exists(Name)**

*express*   一个代表 **Bookmarks** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明** |
| -------- | ------------- | ------------ | -------- |
| *Name*   | 必选          | **String**   | 书签名   |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 如果存在名为 "temp" 的书签，则选中 */ function test() { 	let bks = Application.ActiveDocument.Bookmarks; 	if(bks.Exists("temp") === true) { 		bks.Item("temp").Select(); 	} }` |

#### **Bookmarks.Item**

返回集合中的单个 **Bookmark** 对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **Bookmarks** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Index*  | 必选          | **Variant**  | 要返回的单个对象。可以是代表序号位置的 Number 类型值，或代表单个对象名称的 String 类型值。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 选中名为"temp"的书签 */ Application.ActiveDocument.Bookmarks.Item("temp").Select()` |

**成员属性**

#### **Bookmarks.Application**

返回一个 [Application ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象代表 WPS 应用程序。

**语法**

**express.Application**

*express*   一个代表 **Bookmarks** 对象的变量。

#### **Bookmarks.Count**

返回 **Bookmarks** 集合中的项目数。**Number** 类型，只读。

**语法**

**express.Count**

*express*   一个代表 **Bookmarks** 对象的变量。

#### **Bookmarks.Creator**

返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。**Long** 类型，只读。

**语法**

**express.Creator**

*express*   一个代表 **Bookmarks** 对象的变量。

**说明**

如果对象是在 WPS 中创建的，则 **Creator** 属性返回十六进制数 4D535744，代表字符串“WPS”。该属性主要是为 Macintosh 机的应用设计的，在该机上每个应用程序都有一个四字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参考 WPS OfficeMacintosh Edition 中的语言参考帮助。

#### **Bookmarks.DefaultSorting**

返回或设置**“插入”**菜单上**“书签”**对话框中显示的书签名的排序依据选项。**WdBookmarkSortBy** 类型，可读写。

**语法**

**express.DefaultSorting**

*express*   一个代表 **Bookmarks** 对象的变量。

**说明**

该属性不影响 **Bookmark** 对象在 **Bookmarks** 集合中的顺序。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例按位置对书签进行排序，然后显示“书签”对话框。*/ function test() {     Application.ActiveDocument.Bookmarks.DefaultSorting = wdSortByLocation     Dialogs.Item(wdDialogInsertBookmark).Show() }` |

#### **Bookmarks.Parent**

返回一个 **Object** 类型的值，该值代表指定 **Bookmarks** 集合中的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Bookmarks** 对象的变量。

#### **Bookmarks.ShowHidden**

如果 **Bookmarks** 集合中包括隐藏书签，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ShowHidden**

*express*   一个代表 **Bookmarks** 对象的变量。

**说明**

**ShowHidden** 属性还可控制是否在**“插入”**菜单的**“书签”**对话框中列出隐藏书签。在文档中插入交叉引用时，会自动插入隐藏书签。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 显示“书签”对话框，并设置显示隐藏书签 */ function test() { 	Application.ActiveDocument.Bookmarks.ShowHidden = false; 	Application.Dialogs.Item(wdDialogInsertBookmark).Show(); }` |

适用环境：web

适用平台：windows/linux