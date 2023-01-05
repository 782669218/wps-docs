#### **Reviewer**



代表已修订过的文档的单个审阅者。 Reviewer 对象是 **Reviewers** 集合的一个成员。

**说明**

使用 Reviewers (*索引*) （其中 *index* 是审阅者的名称或编号）可返回 Reviewer 对象。 使用 Visible 属性可以显示或隐藏文档中的单个审阅者。 

**属性**

|                                                              | 名称        | 说明               |
| ------------------------------------------------------------ | ----------- | ------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Visible** | 指定的对象是否可见 |

**成员属性**

#### **Reviewer.Visible**

指定的对象是否可见

**语法**

**express.Visible**

*express*   一个代表 **Reviewer** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 下面的代码示例隐藏名为"zhangsan"的审阅者，并显示名为"lisi"的审阅者。 这假定"zhangsan"和"lisi"是 Reviewers 集合 的成员。 如果没有这两个审阅者，您将收到错误。 */ function test() {  	let vw = Application.ActiveWindow.View   	if (vw.Reviewers.Count < 2)  		return   	vw.Reviewers.Item("chongge").Visible = false   	vw.Reviewers("lisi").Visible = true  }` |

适用环境：web

适用平台：windows/linux