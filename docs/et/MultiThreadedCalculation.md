**MultiThreadedCalculation**



返回或设置并发计算模式。

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Enabled**     | 使用 **Enabled** 属性可以在运行时启用或禁用 **MultiThreadedCalculation** 对象。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ThreadCount** | 获取进程的线程总数，这些线程是指定的 **MultiThreadedCalculation** 对象的一部分。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ThreadMode**  | 返回或设置指定的 **MultiThreadedCalculation** 对象的线程模式。可读/写 **XlThreadMode** 类型。 |

**成员属性**

#### **MultiThreadedCalculation.Application**

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **MultiThreadedCalculation** 对象的变量。

#### **MultiThreadedCalculation.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **MultiThreadedCalculation** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **MultiThreadedCalculation.Enabled**

使用 **Enabled** 属性可以在运行时启用或禁用 **MultiThreadedCalculation** 对象。可读/写。

**语法**

**express.Enabled**

*express*   一个代表 **MultiThreadedCalculation** 对象的变量。

#### **MultiThreadedCalculation.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **MultiThreadedCalculation** 对象的变量。

#### **MultiThreadedCalculation.ThreadCount**

获取进程的线程总数，这些线程是指定的 **MultiThreadedCalculation** 对象的一部分。

**语法**

**express.ThreadCount**

*express*   一个代表 **MultiThreadedCalculation** 对象的变量。

#### **MultiThreadedCalculation.ThreadMode**

返回或设置指定的 **MultiThreadedCalculation** 对象的线程模式。可读/写 **XlThreadMode** 类型。

**语法**

**express.ThreadMode**

*express*   一个代表 **MultiThreadedCalculation** 对象的变量。

**说明**

线程模式可以设置为 **XlThreadModeAutomatic** 或 **XlThreadModeManual**。

适用环境：web

适用平台：windows/linux