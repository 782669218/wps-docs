#### **ParagraphFormat2**



代表文本范围的段落格式。

**说明**

以下示例将活动 PowerPoint 演示文稿第一张幻灯片第二个形状中的段落左对齐。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `Application.ActiveWorkbook.ActiveChart.Shapes.Item(1).SmartArt.AllNodes.Item(1).TextFrame2.TextRange.Text="Node 1" ` |

**属性**

|                                                              | 名称                      | 说明                                                         |
| ------------------------------------------------------------ | ------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Alignment**             | 获取或设置一个值，该值指定段落的对齐方式。可读写。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**           | 获取一个代表包含该对象的应用程序的对象。只读。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BaseLineAlignment**     | 获取或设置一个常量，该常量代表段落中字体的垂直位置。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Bullet**                | 获取段落的 **BulletFormat2** 对象。只读。                    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**               | 获取一个值，该值代表创建 **ParagraphFormat2** 对象的应用程序。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FarEastLineBreakLevel** | 获取或设置指定段落的东亚换行符控制级别。可读写。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FirstLineIndent**       | 获取或设置首行缩进或悬挂缩进的值（以磅值表示）。可读写。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HangingPunctuation**    | 确定指定段落中的标点是否可以溢出边界。可读写。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IndentLevel**           | 获取或设置一个值，该值代表分配给选定段落中的文本的缩进级别。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LeftIndent**            | 返回或设置一个值，该值代表指定段落的左缩进值（以磅为单位）。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LineRuleAfter**         | 确定是否将每段最后一行后面的行距设为特定的磅数或行数。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LineRuleBefore**        | 确定是否将每段首行前面的行距设为特定的磅数或行数。可读写。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LineRuleWithin**        | 确定是否将基线间的行距设为特定的磅数或行数。可读写。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                | 获取 **ParagraphFormat2** 对象的父对象。只读。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RightIndent**           | 获取或设置指定段落的右缩进量（以磅为单位）。可读写。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SpaceAfter**            | 获取或设置指定段落的段后间距（以磅为单位）。可读写。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SpaceBefore**           | 获取或设置指定段落的段前间距（以磅为单位）。可读写。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SpaceWithin**           | 获取或设置指定段落中基准行之间的距离（以磅或行为单位）。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TabStops**              | 获取一个代表指定段落的所有自定义制表位的 **TabStops2** 集合。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextDirection**         | 获取或设置指定段落的文本方向。可读写。                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WordWrap**              | 确定应用程序是否在指定段落的语句中对拉丁语文本换行。可读写。 |

**成员属性**

#### **ParagraphFormat2.Alignment**

获取或设置一个值，该值指定段落的对齐方式。可读写。

**语法**

**express.Alignment**

*express*   一个代表 **ParagraphFormat2** 对象的变量。

#### **ParagraphFormat2.Application**

获取一个代表包含该对象的应用程序的对象。只读。

**语法**

**express.Application**

*express*   一个代表 **ParagraphFormat2** 对象的变量。

#### **ParagraphFormat2.BaseLineAlignment**

获取或设置一个常量，该常量代表段落中字体的垂直位置。可读写。

**语法**

**express.BaseLineAlignment**

*express*   一个代表 **ParagraphFormat2** 对象的变量。

#### **ParagraphFormat2.Bullet**

获取段落的 **BulletFormat2** 对象。只读。

**语法**

**express.Bullet**

*express*   一个代表 **ParagraphFormat2** 对象的变量。

#### **ParagraphFormat2.Creator**

获取一个值，该值代表创建 **ParagraphFormat2** 对象的应用程序。只读。

**语法**

**express.Creator**

*express*   一个代表 **ParagraphFormat2** 对象的变量。

#### **ParagraphFormat2.FarEastLineBreakLevel**

获取或设置指定段落的东亚换行符控制级别。可读写。

**语法**

**express.FarEastLineBreakLevel**

*express*   一个代表 **ParagraphFormat2** 对象的变量。

#### **ParagraphFormat2.FirstLineIndent**

 获取或设置首行缩进或悬挂缩进的值（以磅值表示）。可读写。

**语法**

**express.FirstLineIndent**

*express*   一个代表 **ParagraphFormat2** 对象的变量。

#### **ParagraphFormat2.HangingPunctuation**

确定指定段落中的标点是否可以溢出边界。可读写。

**语法**

**express.HangingPunctuation**

*express*   一个代表 **ParagraphFormat2** 对象的变量。

#### **ParagraphFormat2.IndentLevel**

获取或设置一个值，该值代表分配给选定段落中的文本的缩进级别。可读写。

**语法**

**express.IndentLevel**

*express*   一个代表 **ParagraphFormat2** 对象的变量。

#### **ParagraphFormat2.LeftIndent**

返回或设置一个值，该值代表指定段落的左缩进值（以磅为单位）。可读写。

**语法**

**express.LeftIndent**

*express*   一个代表 **ParagraphFormat2** 对象的变量。

#### **ParagraphFormat2.LineRuleAfter**

确定是否将每段最后一行后面的行距设为特定的磅数或行数。可读写。

**语法**

**express.LineRuleAfter**

*express*   一个代表 **ParagraphFormat2** 对象的变量。

#### **ParagraphFormat2.LineRuleBefore**

确定是否将每段首行前面的行距设为特定的磅数或行数。可读写。

**语法**

**express.LineRuleBefore**

*express*   一个代表 **ParagraphFormat2** 对象的变量。

#### **ParagraphFormat2.LineRuleWithin**

确定是否将基线间的行距设为特定的磅数或行数。可读写。

**语法**

**express.LineRuleWithin**

*express*   一个代表 **ParagraphFormat2** 对象的变量。

#### **ParagraphFormat2.Parent**

获取 **ParagraphFormat2** 对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **ParagraphFormat2** 对象的变量。

#### **ParagraphFormat2.RightIndent**

获取或设置指定段落的右缩进量（以磅为单位）。可读写。

**语法**

**express.RightIndent**

*express*   一个代表 **ParagraphFormat2** 对象的变量。

#### **ParagraphFormat2.SpaceAfter**

获取或设置指定段落的段后间距（以磅为单位）。可读写。

**语法**

**express.SpaceAfter**

*express*   一个代表 **ParagraphFormat2** 对象的变量。

#### **ParagraphFormat2.SpaceBefore**

获取或设置指定段落的段前间距（以磅为单位）。可读写。

**语法**

**express.SpaceBefore**

*express*   一个代表 **ParagraphFormat2** 对象的变量。

#### **ParagraphFormat2.SpaceWithin**

获取或设置指定段落中基准行之间的距离（以磅或行为单位）。可读写。

**语法**

**express.SpaceWithin**

*express*   一个代表 **ParagraphFormat2** 对象的变量。

#### **ParagraphFormat2.TabStops**

获取一个代表指定段落的所有自定义制表位的 **TabStops2** 集合。只读。

**语法**

**express.TabStops**

*express*   一个代表 **ParagraphFormat2** 对象的变量。

#### **ParagraphFormat2.TextDirection**

获取或设置指定段落的文本方向。可读写。

**语法**

**express.TextDirection**

*express*   一个代表 **ParagraphFormat2** 对象的变量。

#### **ParagraphFormat2.WordWrap**

确定应用程序是否在指定段落的语句中对拉丁语文本换行。可读写。

**语法**

**express.WordWrap**

*express*   一个代表 **ParagraphFormat2** 对象的变量。

适用环境：web

适用平台：windows/linux