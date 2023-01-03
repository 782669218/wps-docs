#### **Sync**



WPS 中的 **Document** 对象、ET 中的 **Workbook** 对象以及 WPP 中的 **Presentation** 对象的 **Sync** 属性均可返回一个 **Sync** 对象。

| ![img]()注释                                                |
| ----------------------------------------------------------- |
| 自 WPS Office 2015 开始，此对象或成员已弃用，不应再行使用。 |

**说明**

使用 **Sync** 对象可管理存储在 SharePoint网站中的共享文档的本地副本和服务器副本之间的同步。**Status** 属性返回有关同步当前状态的重要信息。使用 **GetUpdate** 方法可刷新同步状态。使用 **LastSyncTime**、**ErrorType** 和 **WorkspaceLastChangedBy** 属性可返回其他信息。

请参阅 **Status** 属性以获取有关共享文档的本地副本和服务器副本之间的差异和冲突的其他信息。

使用 **PutUpdate** 方法可将本地更改保存到服务器中。在没有进行本地更改时，关闭并重新打开文档可检索来自服务器的最新版本。使用 **ResolveConflict** 方法可解决本地副本和服务器副本之间的差异，或者使用 **OpenVersion** 方法可在当前打开的文档的本地版本旁打开另一版本。

由于 **Sync** 对象的 **GetUpdate**、**PutUpdate** 和 **ResolveConflict** 方法以异步方式完成任务，因此它们不返回状态代码。**Sync** 对象通过单个事件提供重要的状态信息，开发人员可以通过以下应用程序特定事件来访问这些信息：

- 在 WPS 中，通过 **Document** 对象的 **Sync** 事件或 **Application** 对象的 **DocumentSync** 事件；
- 在 ET 中，通过 **Workbook** 对象的 **Sync** 事件或 **Application** 对象的 **WorkbookSync** 事件；
- 在 WPP 中，通过 **Application** 对象的 **PresentationSync** 事件。

上述 **Sync** 事件返回一个 **msoSyncEventType** 值。

无论活动文档是启用还是禁用了共享和同步，**Sync** 对象模型都可用。当活动文档未共享或没有启用同步时，**Document**、**Workbook** 和 **Presentation** 对象的 **Sync** 属性不返回 **Nothing**。使用 **Status** 属性可确定文档是否共享以及是否启用了同步。

并不是所有文档同步问题都会引发可捕获的运行时错误。在使用了 **Sync** 对象的方法后，最好对 **Status** 属性进行检查。如果 **Status** 属性是 **msoSyncStatusError**，则应检查 **ErrorType** 属性，以获取有关已发生的错误类型的其他信息。

在许多情况下，解决错误条件的最好方法是调用 **GetUpdate** 方法。例如，如果对 **PutUpdate** 的调用导致错误条件的产生，则对 **GetUpdate** 的调用会将状态重置为 **msoSyncStatusLocalChanges**。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的示例将基于活动文档的状态说明 Sync 对象的多种方法。*/ function test(){ 	var objSync = Application.ActiveWorkbook.Sync; 	if(objSync.Status > msoSyncEventDownloadFailed) 	{ 		switch (objSync.Status) 		{ 			case msoSyncStatusConflict: 				objSync.ResolveConflict(msoSyncConflictMerge);             	ActiveDocument.Save();             	objSync.ResolveConflict(msoSyncConflictClientWins);             	strStatus = "Conflict resolved by merging changes.";             	break; 			case msoSyncStatusError: 				strStatus = "Last error type: " + objSync.ErrorType; 	            break; 			case msoSyncStatusLatest: 				strStatus = "Document copies already in sync.";             	break; 			case msoSyncStatusLocalChanges: 				objSync.PutUpdate();             	strStatus = "Local changes saved to server.";             	break; 			case msoSyncStatusNewerAvailable: 				objSync.GetUpdate();             	strStatus = "Local copy updated from server.";             	break; 			case msoSyncStatusSuspended: 				objSync.Unsuspend();             	strStatus = "Synchronization resumed.";             	break; 		} 	} 	else 		strStatus = "Not a shared workspace document."; 	alert(strStatus + " -- Sync Information"); }` |

**方法**

|                                                              | 名称                | 说明                                               |
| ------------------------------------------------------------ | ------------------- | -------------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **OpenVersion**     | 在当前打开的本地版本旁打开共享文档的另一个版本。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PutUpdate**       | 用共享文档的本地副本更新其服务器副本。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ResolveConflict** | 解决共享文档的本地副本和服务器副本之间的冲突。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Unsuspend**       | 继续执行共享文档的本地副本和服务器副本之间的同步。 |

**属性**

|                                                              | 名称                       | 说明                                                         |
| ------------------------------------------------------------ | -------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**            | 获取一个 **Application** 对象，代表 **Sync** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                | 获取一个 32 位整数，指示创建 **Sync** 对象时所使用的应用程序。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ErrorType**              | 获取指示最新文档同步错误类型的 **MsoSyncErrorType** 常量。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LastSyncTime**           | 获取上一次将活动文档的本地副本与服务器副本同步的日期和时间。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                 | 获取 **Sync** 对象的 **Parent** 对象。只读。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Status**                 | 获取活动文档的本地副本与服务器副本的同步状态。只读。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WorkspaceLastChangedBy** | 显示最后一个在共享文档服务器副本中保存更改的用户的显示名称。只读。 |

**成员方法**

#### **Sync.OpenVersion**

在当前打开的本地版本旁打开共享文档的另一个版本。

**语法**

**express.OpenVersion(SyncVersionType)**

*express*   一个代表 **Sync** 对象的变量。

**参数**

| **名称**          | **必选/可选** | **数据类型**           | **说明**         |
| ----------------- | ------------- | ---------------------- | ---------------- |
| *SyncVersionType* | 必选          | **MsoSyncVersionType** | 代表版本的类型。 |

**说明**

使用 **OpenVersion** 方法可在当前打开的本地版本旁打开上次查看的共享文档版本 (**msoSyncVersionLastViewed**) 或其服务器副本 (**msoSyncVersionServer**)。

**The msoSyncVersionLastViewed** 选项显示用户用服务器副本覆盖本地副本时创建的文档副本。例如，如果使用 **msoSyncConflictServerWins** 选项调用 **ResolveConflict** 方法，则将保存您的本地更改，并可以通过调用 **OpenVersion(msoSyncVersionLastViewed)** 进行查看。

并非所有文档同步问题都会引发可捕获的运行时错误。在使用 **Sync** 对象执行完一个操作后，最好对 **Status** 属性进行检查；如果 **Status** 属性是 **msoSyncStatusError**，则应检查 **ErrorType** 属性，以获取有关已发生的错误类型的其他信息。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的示例将在当前打开的共享文档本地版本旁打开该文档的服务器副本。*/ function test(){ 	var objSync = Application.ActiveWorkbook.Sync;  	objSync.OpenVersion(Application.Enum.msoSyncVersionServer); }` |

#### **Sync.PutUpdate**

用共享文档的本地副本更新其服务器副本。

**语法**

**express.PutUpdate()**

*express*   一个代表 **Sync** 对象的变量。

**说明**

如果客户端未识别出最近对共享文档服务器副本所做的更改，**PutUpdate** 方法可能会遇到冲突。在调用 **PutUpdate** 之前先调用 **GetUpdate** 方法可刷新服务器副本的状态并检测可能的冲突。

如果本地文档有未保存的更改，**PutUpdate** 方法将发生运行时错误。

并非所有文档同步问题都会引发可捕获的运行时错误。在使用 **Sync** 对象执行完一个操作后，最好对 **Status** 属性进行检查；如果 **Status** 属性是 **msoSyncStatusError**，则应检查 **ErrorType** 属性，以获取有关已发生的错误类型的其他信息。

在许多情况下，解决错误条件的最好方法是调用 **GetUpdate** 方法。例如，如果对 **PutUpdate** 的调用导致错误条件的产生，则对 **GetUpdate** 的调用会将状态重置为 **msoSyncStatusLocalChanges**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果文档的本地副本已被编辑，则下面的示例会使用 PutUpdate 方法根据本地副本更新服务器副本。*/ function test(){ 	var objSync = Application.ActiveWorkbook.Sync; 	if(objSync.Status == Application.Enum.msoSyncStatusLocalChanges) {     	objSync.PutUpdate();     	strStatus = "Local changes saved to server.";     	alert(strStatus + "Sync Information"); 	} }` |

#### **Sync.ResolveConflict**

解决共享文档的本地副本和服务器副本之间的冲突。

**语法**

**express.ResolveConflict(SyncConflictResolution)**

*express*   一个代表 **Sync** 对象的变量。

**参数**

| **名称**                 | **必选/可选** | **数据类型**                      | **说明**             |
| ------------------------ | ------------- | --------------------------------- | -------------------- |
| *SyncConflictResolution* | 必选          | **MsoSyncConflictResolutionType** | 指定解决冲突的方式。 |

**说明**

使用 **ResolveConflict** 方法可解决活动文档的本地副本和服务器副本之间的差异。使用 **msoSyncConflictMerge** 选项（不适用于 ET 工作簿）可以合并各个文档的更改。通过 **msoSyncConflictClientWins** 选项使用本地更改替换服务器副本，或者通过 **msoSyncConflictServerWins** 选项使用更改的服务器副本替换本地副本。

**msoSyncConflictMerge** 选项将对服务器副本所做的更改合并到本地副本中，但实际上并没有解决冲突。要解决冲突并成功地合并更改，必须在合并更改后保存活动文档，然后再次使用 **msoSyncConflictClientWins** 选项调用 **ResolveConflict** 方法。

如果客户端未检测出对共享文档服务器副本的最新更改，**ResolveConflict** 方法可能会遇到冲突条件。在调用 **ResolveConflict** 之前先调用 **GetUpdate** 方法可刷新服务器副本的状态并检测可能的冲突。

如果本地文档有未保存的更改或者在文档的两个副本之间不存在冲突，**ResolveConflict** 方法就会发生运行时错误。

并不是所有文档同步问题都会引发可捕获的运行时错误。在使用 **Sync** 对象执行完一个操作后，最好对 **ErrorType** 属性进行检查。如果 **Status** 属性是 **msoSyncStatusError**，则应检查 **Status** 属性，以进一步获取有关所发生错误类型的信息。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的示例试图通过合并更改来解决活动文档的本地副本和服务器副本之间的冲突。*/ function test(){ 	var objSync = Application.ActiveWorkbook.Sync; 	if(objSync.Status == Application.Enum.msoSyncStatusConflict) {     	objSync.ResolveConflict(Application.Enum.msoSyncConflictMerge);     	Application.ActiveWorkbook.Save();     	objSync.ResolveConflict(Application.Enum.msoSyncConflictClientWins);     	strStatus = "Conflict resolved by merging changes.";     	alert(strStatus + "Sync Information"); 	} }` |

#### **Sync.Unsuspend**

继续执行共享文档的本地副本和服务器副本之间的同步。

**语法**

**express.Unsuspend()**

*express*   一个代表 **Sync** 对象的变量。

**说明**

使用 **Unsuspend** 方法可在 **Status** 属性返回 **msoSyncStatusSuspended** 时继续执行文档同步。

并非所有文档同步问题都会引发可捕获的运行时错误。在使用 **Sync** 对象执行完一个操作后，最好对 **Status** 属性进行检查；如果 **Status** 属性是 **msoSyncStatusError**，则应检查 **ErrorType** 属性，以获取有关已发生的错误类型的其他信息。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的示例将在文档同步暂停后继续执行文档同步。*/ function test(){ 	var objSync = Application.ActiveWorkbook.Sync; 	if(objSync.Status == Application.Enum.msoSyncStatusSuspended) {     	objSync.Unsuspend();     	alert(strStatus + "Sync Information"); 	} }` |

**成员属性**

#### **Sync.Application**

获取一个 **Application** 对象，代表 **Sync** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

*express*   一个代表 **Sync** 对象的变量。

#### **Sync.Creator**

获取一个 32 位整数，指示创建 **Sync** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

*express*   一个代表 **Sync** 对象的变量。

#### **Sync.ErrorType**

获取指示最新文档同步错误类型的 **MsoSyncErrorType** 常量。只读。

**语法**

**express.ErrorType**

*express*   一个代表 **Sync** 对象的变量。

**说明**

使用 **ErrorType** 属性确定最新文档同步错误的类型。并不是所有文档同步问题都会引发可捕获的运行时错误。在使用 **Sync** 对象执行完一项操作后，最好对 **Status** 属性进行检查；如果 **Status** 属性是 **msoSyncStatusError**，则应检查 **ErrorType** 属性，以获取所发生错误类型的其他信息。

#### **Sync.LastSyncTime**

获取上一次将活动文档的本地副本与服务器副本同步的日期和时间。只读。

**语法**

**express.LastSyncTime**

*express*   一个代表 **Sync** 对象的变量。

**说明**

使用 **LastSyncTime** 属性可确定距活动文档的本地副本上一次与服务器副本同步已有多长时间。检查 **Status** 属性可确定本地副本与服务器副本是否同步。

如果未将活动文档配置为在本地副本与服务器副本之间进行同步，**LastSyncTime** 属性就会产生运行时错误。

**示例**

#### **Sync.Parent**

获取 **Sync** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **Sync** 对象的变量。

#### **Sync.Status**

获取活动文档的本地副本与服务器副本的同步状态。只读。

**语法**

**express.Status**

*express*   一个代表 **Sync** 对象的变量。

**说明**

使用 **Status** 属性可确定活动文档的本地副本是否与共享服务器副本同步。使用 **GetUpdate** 方法可刷新状态。对于不同的状态条件，请分别使用下列方法和属性：

- **msoSyncStatusConflict** - 当本地副本和服务器副本都发生更改时为 **True**。使用 **ResolveConflict** 方法可解决这些差异。
- **msoSyncStatusError** - 检查 **ErrorType** 属性。
- **msoSyncStatusLocalChanges** - 当只有本地副本发生更改时为 **True**。使用 **PutUpdate** 方法可将本地更改保存到服务器副本中。
- **msoSyncStatusNewerAvailable** - 当只有服务器副本发生更改时为 **True**。关闭并重新打开文档可使用来自服务器中的最新副本。
- **msoSyncStatusSuspended** - 使用 **Unsuspend** 方法可恢复同步。

**Status** 属性按照下面的优先级顺序从列表中返回单个常量：

1. **msoSyncStatusNoSharedWorkspace**
2. **msoSyncStatusError**
3. **msoSyncStatusSuspended**
4. **msoSyncStatusConflict**
5. **msoSyncStatusNewerAvailable**
6. **msoSyncStatusLocalChanges**
7. **msoSyncStatusLatest**

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的示例将检查 Status 属性，并在必要时进行适当的操作以使文档的本地副本与服务器副本同步。*/ function test(){ 	var objSync = Application.ActiveWorkbook.Sync; 	if(objSync.Status > Application.Enum.msoSyncEventDownloadFailed) 	{ 		switch (objSync.Status) 		{ 			case Application.Enum.msoSyncStatusConflict: 				objSync.ResolveConflict(Application.Enum.msoSyncConflictMerge);             	　　　　　　　　 Application.ActiveDocument.Save();             	　　　　　　　　 objSync.ResolveConflict(Application.Enum.msoSyncConflictClientWins);             	　　　　　　　　 strStatus = "Conflict resolved by merging changes.";             	break; 			case Application.Enum.msoSyncStatusError: 				strStatus = "Last error type: " + objSync.ErrorType; 	           　　　　　　  break; 			case Application.Enum.msoSyncStatusLatest: 				strStatus = "Document copies already in sync.";             	　　　　　　　　 break; 			case Application.Enum.msoSyncStatusLocalChanges: 				objSync.PutUpdate();             	　　　　　　　　 strStatus = "Local changes saved to server.";             	　　　　　　　　 break; 			case Application.Enum.msoSyncStatusNewerAvailable: 				objSync.GetUpdate();             	　　　　　　　　 strStatus = "Local copy updated from server.";             	　　　　　　　　 break; 			case Application.Enum.msoSyncStatusSuspended: 				objSync.Unsuspend();             	　　　　　　　　 strStatus = "Synchronization resumed.";             	　　　　　　　　 break; 		} 	} 	else 	　　　　strStatus = "Not a shared workspace document."; 	　　　　alert(strStatus + "Sync Information"); }` |

#### **Sync.WorkspaceLastChangedBy**

显示最后一个在共享文档服务器副本中保存更改的用户的显示名称。只读。

**语法**

**express.WorkspaceLastChangedBy**

*express*   一个代表 **Sync** 对象的变量。

**说明**

如果未将活动文档配置为在本地副本和服务器副本之间进行同步，**WorkspaceLastChangedBy** 属性就会产生运行时错误。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的示例将检查共享文档的本地副本和服务器副本之间的冲突，并报告最后一个在服务器副本中保存更改的用户的名称。*/ function test(){ 	var objSync = Application.ActiveWorkbook.Sync; 	if(objSync.Status == Application.Enum.msoSyncStatusConflict) {     	strStatus = "The server copy has been changed by" + objSync.WorkspaceLastChangedBy;     	alert(strStatus + "Server Copy Changed"); 	} }` |

适用环境：web

适用平台：windows/linux