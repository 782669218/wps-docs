#### **ListGallery**



代表单个列表格式库。**ListGallery** 对象是 **ListGalleries** 集合的一个成员。

**说明**

每个 **ListGallery** 对象代表**“项目符号和编号”**对话框中的三个选项卡之一。

使用 **ListGalleries**(*Index*) 可返回一个 **ListGallery** 对象，其中 *Index* 为 **wdBulletGallery**、**wdNumberGallery** 或 **wdOutlineNumberGallery**。

以下示例返回**“项目符号和编号”**对话框中**“项目符号”**选项卡上的第三种列表格式（不包括**“无”**），并将其应用于所选内容。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { 	let temp3 = Application.ListGalleries.Item(wdBulletGallery).ListTemplates.Item(3) 	Application.Selection.Range.ListFormat.ApplyListTemplate(temp3) }` |

 

ListGallery

 

 

Modified

 

 

ListGallery

 

 

Reset

 

**属性**

|                                                              | 名称              | 说明                                                         |
| ------------------------------------------------------------ | ----------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ListTemplates** | 返回一个 **ListTemplates** 集合，该集合代表指定列表库的所有列表格式。只读。 |

**成员属性**

#### **ListGallery.ListTemplates**

返回一个 **ListTemplates** 集合，该集合代表指定列表库的所有列表格式。只读。

**语法**

**express.ListTemplates**

*express*   一个代表 **ListGallery** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 以下代码将选区设置为项目符号的第3钟符号模板 */ function test() { 	let temp3 = Application.ListGalleries.Item(wdBulletGallery).ListTemplates.Item(3) 	Application.Selection.Range.ListFormat.ApplyListTemplate(temp3) }` |

适用环境：web

适用平台：windows/linux