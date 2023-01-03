#### **PageNumber**



代表页眉或页脚中的页码。**PageNumber** 对象是 **PageNumbers** 集合的一个成员。**PageNumbers** 集合包含单个页眉或页脚中的所有页码。

**说明**

使用 **PageNumbers**(*Index*) 可返回一个 **PageNumber** 对象，其中 *Index* 为索引号。在大多数情况下，一个页眉或页脚只包含一个页码（索引号为 1）。以下示例将活动文档第一节的主页眉中的起始页码居中。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActiveDocument.Sections.Item(1).Headers(wdHeaderFooterPrimary).PageNumbers.Item(1).Alignment = wdAlignPageNumberCenter` |

 

 

Add

 

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let sect = Application.ActiveDocument.Sections.Item(1)     sect.Footers(wdHeaderFooterPrimary).PageNumbers.Add(wdAlignPageNumberLeft, true) }` |

**属性**

|                                                              | 名称          | 说明                                                         |
| ------------------------------------------------------------ | ------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Alignment** | 返回或设置一个 **WdPageNumberAlignment** 常量，该常量代表页码的对齐方式，可读写。 |

**成员属性**

#### **PageNumber.Alignment**

返回或设置一个 **WdPageNumberAlignment** 常量，该常量代表页码的对齐方式，可读写。

**语法**

**express.Alignment**

*express*   一个代表 **PageNumber** 对象的变量。

适用环境：web

适用平台：windows/linux