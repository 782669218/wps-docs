#### **SmartDocument**



WPS 中的 **Document** 对象和 ET 中的 **Workbook** 对象的 **SmartDocument** 属性返回一个 **SmartDocument** 对象。

**说明**

使用 **SmartDocument** 对象可管理附加到活动文档上的 XML 扩展包。

使用 **SmartDocument** 对象的 **SolutionID** 和 **SolutionURI** 属性可检索有关附加到活动文档或工作簿上的 XML 扩展包的信息。使用 **PickSolution** 方法可允许用户从列表中选择可用的 XML 扩展包以便附加到活动文档或工作簿中。使用 **RefreshPane** 方法可刷新智能文档的**“文档操作”**任务窗格。

无论文档是否附加了 XML 扩展包，**SmartDocument** 对象模型均可用。当活动文档没有附加 XML 扩展包时，**Document** 或 **Workbook** 对象的 **SmartDocument** 属性不返回 **Nothing**。检查 **SolutionID** 属性可确定活动文档是否附加了 XML 扩展包。

**方法**

|                                                              | 名称             | 说明                                                         |
| ------------------------------------------------------------ | ---------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PickSolution** | 显示一个对话框，允许用户选择可用的 XML 扩展包，以便将其附加到 WPS 活动文档或 ET 工作簿中。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **RefreshPane**  | 为 WPS 中的活动文档或 ET 中的工作簿刷新**“文档操作”**任务窗格。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 获取一个 **Application** 对象，代表 **SmartDocument** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 获取一个 32 位整数，指示创建 **SmartDocument** 对象时所使用的应用程序。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SolutionID**  | 获取或设置 ID，通常为全局唯一标识符 (GUID)，它标识附加在活动的 WPS 文档或者 ET 工作簿上的 XML 扩展包。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SolutionURL** | 获取或设置一个绝对 URL，该 URL 提供指向附加到活动的 WPS 文档或者 ET 工作簿上的 XML 扩展包文件的完整路径。可读写。 |

**成员方法**

#### **SmartDocument.PickSolution**

显示一个对话框，允许用户选择可用的 XML 扩展包，以便将其附加到 WPS 活动文档或 ET 工作簿中。

**语法**

**express.PickSolution(ConsiderAllSchemas )**

*express*   一个代表 **SmartDocument** 对象的变量。

**参数**

| **名称**             | **必选/可选** | **数据类型** | **说明**                                                     |
| -------------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *ConsiderAllSchemas* | 必选          | **Boolean**  | 如果为 True，则会显示用户计算机上安装的所有可用 XML 扩展包。如果为 False，则会仅显示可用于活动文档的 XML 扩展包。默认值为 False。 |

**说明**

使用 **PickSolution** 方法可允许用户从列表中选择 XML 扩展包。附加在活动文档或工作簿上的架构将确定可用的 XML 扩展包。

**PickSolution** 方法不会返回一个值以表明用户是选择了 XML 扩展包还是单击了对话框中的 **Cancel**。在调用 **SolutionID** 后，检查 **PickSolution** 属性可确定是否已附加了 XML 扩展包。

如果智能文档的开发人员未能在 XML 扩展包清单文件中指定“targetApplication”，则 **PickSolution** 所显示的列表中可能会包含未面向活动应用程序的 XML 扩展包。例如，ET 用户可能会看到专门面向 WPS 的 XML 扩展包。在这些情况下，用户可能会选择不适合活动应用程序的 XML 扩展包。

关于智能文档或智能文档 XML 扩展包的详细信息，请参阅WPS开发网站上的 Smart Document Software Development Kit (SDK)。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的示例将检查 SolutionID 属性以确定活动 WPS 文档是否已有附加的 XML 扩展包。如果没有，它会显示一个对话框，允许用户选择可用的 XML 扩展包，然后显示智能文档的属性。*/ function test(){ 	var objSmartDoc = Application.ActiveWorkbook.SmartDocument; 	if(objSmartDoc.SolutionID == "None" || objSmartDoc.SolutionID == "")     	objSmartDoc.PickSolution(true); 	if(objSmartDoc.SolutionID > "None" && objSmartDoc.SolutionID > "")  	{     	　　　　let strSmartDocInfo = "SolutionID: " + objSmartDoc.SolutionID + "\r\n" + "SolutionURL: " + objSmartDoc.SolutionURL;     	　　　　alert(strSmartDocInfo + "Smart Doc Properties"); 	} 	else      	　　　　alert("The user clicked Cancel."); }` |

#### **SmartDocument.RefreshPane**

为 WPS 中的活动文档或 ET 中的工作簿刷新**“文档操作”**任务窗格。

**语法**

**express.RefreshPane()**

*express*   一个代表 **SmartDocument** 对象的变量。

**说明**

如果活动文档没有附加 XML 扩展包，**RefreshPane** 方法就会产生错误。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的示例确定活动的 ET 工作簿是否附加了 XML 扩展包。如果是，它将刷新智能文档的“文档操作”任务窗格。*/ function test(){ 	var objSmartDoc = Application.ActiveWorkbook.SmartDocument; 	if(objSmartDoc.SolutionID > "None")      	objSmartDoc.RefreshPane; 	else 		alert("No XML expansion pack attached."); }` |

**成员属性**

#### **SmartDocument.Application**

获取一个 **Application** 对象，代表 **SmartDocument** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

*express*   一个代表 **SmartDocument** 对象的变量。

#### **SmartDocument.Creator**

获取一个 32 位整数，指示创建 **SmartDocument** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

*express*   一个代表 **SmartDocument** 对象的变量。

#### **SmartDocument.SolutionID**

获取或设置 ID，通常为全局唯一标识符 (GUID)，它标识附加在活动的 WPS 文档或者 ET 工作簿上的 XML 扩展包。可读写。

**语法**

**express.SolutionID**

*express*   一个代表 **SmartDocument** 对象的变量。

**说明**

当活动文档没有附加 XML 扩展包时，**SolutionID** 属性返回空字符串或“None”。

通过为 **SolutionID** 和 **SolutionURL** 属性提供适当的值以将可用的 XML 扩展包附加到活动文档上，可以在不使用 **PickSolution** 方法的情况下将该文档转换为智能文档。通过将 **SolutionID** 和 **SolutionUrl** 属性设置为空字符串，可以删除附加的 XML 扩展包。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的示例将通过检查 SolutionID 属性来确定是否在活动 ET 工作簿上附加了 XML 扩展包。*/ function test(){ 	var objSmartDoc = Application.ActiveWorkbook.SmartDocument; 	if(objSmartDoc.SolutionID == "None" || objSmartDoc.SolutionID == "")     	alert("No XML expansion pack attached."); 	else      	alert("Smart document Solution ID: " + objSmartDoc.SolutionID); }` |

#### **SmartDocument.SolutionURL**

获取或设置一个绝对 URL，该 URL 提供指向附加到活动的 WPS 文档或者 ET 工作簿上的 XML 扩展包文件的完整路径。可读写。

**语法**

**express.SolutionURL**

*express*   一个代表 **SmartDocument** 对象的变量。

**说明**

当没有在活动文档上附加 XML 扩展包时，**SolutionUrl** 属性返回一个空字符串。

通过为 **SolutionID** 和 **SolutionUrl** 属性提供适当的值以将可用的 XML 扩展包附加到活动文档上，可以在不使用 **PickSolution** 方法的情况下将该文档转换为智能文档。通过将 **SolutionID** 和 **SolutionUrl** 属性设置为空字符串，可以删除附加的 XML 扩展包。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的示例先确定是否在活动 WPS 文档上附加了 XML 扩展包，然后显示智能文档的解析 URL。*/ function test(){ 	var objSmartDoc = Application.ActiveWorkbook.SmartDocument; 	if(objSmartDoc.SolutionID == "None" || objSmartDoc.SolutionID == "")     	alert("No XML expansion pack attached."); 	else      	alert("Smart document Solution URL: " + objSmartDoc.SolutionURL); }` |

适用环境：web

适用平台：windows/linux