**Page**



代表工作表的页面。使用 **PageSetup**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/PageSetup/PageSetup%20.htm#jsObject_PageSetup)对象及相关方法和属性可通过编程方式定义工作簿的页面布局。

**说明**

使用 **Item** 方法可访问工作簿的特定页面。下面的示例访问活动工作簿的第一页。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `let objPage = Application.ActiveWindow.Panes.Item(1).Pages.Item(1)` |

**属性**

|                                                              | 名称             | 说明                                 |
| ------------------------------------------------------------ | ---------------- | ------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CenterFooter** | 指定要在页脚中居中对齐的图片或文本。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CenterHeader** | 指定要在页眉中居中对齐的图片或文本。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LeftFooter**   | 指定要在页脚中左对齐的图片或文本。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LeftHeader**   | 指定要在页眉中左对齐的图片或文本。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RightFooter**  | 指定要在页脚中右对齐的图片或文本。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RightHeader**  | 指定要在页眉中右对齐的图片或文本。   |

**成员属性**

#### **Page.CenterFooter**

指定要在页脚中居中对齐的图片或文本。

**语法**

**express.CenterFooter**

*express*   一个代表 **Page** 对象的变量。

#### **Page.CenterHeader**

指定要在页眉中居中对齐的图片或文本。

**语法**

**express.CenterHeader**

*express*   一个代表 **Page** 对象的变量。

#### **Page.LeftFooter**

指定要在页脚中左对齐的图片或文本。

**语法**

**express.LeftFooter**

*express*   一个代表 **Page** 对象的变量。

#### **Page.LeftHeader**

指定要在页眉中左对齐的图片或文本。

**语法**

**express.LeftHeader**

*express*   一个代表 **Page** 对象的变量。

#### **Page.RightFooter**

指定要在页脚中右对齐的图片或文本。

**语法**

**express.RightFooter**

*express*   一个代表 **Page** 对象的变量。

#### **Page.RightHeader**

指定要在页眉中右对齐的图片或文本。

**语法**

**express.RightHeader**

*express*   一个代表 **Page** 对象的变量。

适用环境：web

适用平台：windows/linux