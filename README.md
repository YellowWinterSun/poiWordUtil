# poiWordUtil
基于Apache POI封装后的Word文档打印工具类
---------------------------------------------------------------------------------------------------------------------
最近更新：
(1) 2018/12/26 : 更新 图片替换功能。  Now you can use this util to replace picture.

---------------------------------------------------------------------------------------------------------------------

* 特别声明：
    需要用WPS对.docx文档编辑，才能获得最好的支持。本poiWordUtil不支持Microsoft Office的编辑。原因个人分析如下：这可能得追溯到文件byte格式，编码协议等等的原因了吧。office和wps，虽然宏观上我们感知不到其对word文档有啥区别。但是我们即使不了解底层，也会遇到这样的情况，你在office上编辑的word文档格式什么的都很完美，但是发过去给其他人看时，如果对方是用wps或者其他软件打开你的word文档，就会出现格式不一样的问题。在这里，我发现office和wps对于一个XWPFRun的定义不太一样，最直观的感受是，我这一套工具基于wps的规范来定制了。所以用office软件编辑的word文档，在识别XWPFRun（占位符)时，会出错。（如果你有兴趣，可以自行研究一下apache poi的东西，你可以去自行看看XWPFRun在office和wps之间的差异。）
    
* Most Important warning: 
    you have to use WPS(made by China Company) to edit word template. Using Office is not suitable for this git-hub-project and you can't export the word file as you wish.

* DEMO中的docx使用的是WPS编辑的，因此如果用office打开再保存，可能会导致占位符无法正常识别，也就无法进行word打印功能。
请在编辑docx的时，尽量使用WPS。目前apache poi对Office编辑的word文档并不是特别友好，导致一个占位符无法对应一个XWPFRun。
* In this demo, the .docx use by WPS. So if you use office edit demo word file, maybe it will be fail to export. Now I suggest the user use WPS to edit word template, because Apache Poi is not good suitable to Office.(one ${...} cannot match one XWPFRun)

---------------------------------------------------------------------------------------------------------------------
工具：IDEA
框架：Java + Maven 简单的Java项目

---------------------------------------------------------------------------------------------------------------------
注意：当前工具仅仅是本人在项目开发过程中，需要处理大量的word打印工作，在学习Apache Poi后封装的一套工具。有很多代码冗余，还有写的不要的地方，由于开发紧张，所以没有优化，但是使用是没有问题的。大家可以下载代码后，有兴趣自行研究并修改代码，你还可以为工具封装很多有趣的东西。如你可以让每一次替换文本，都可以自己设置字体样式等。

免责申明：仅学习参考使用，请勿用于商业用途。如果您在使用本工具作为商业用途的过程中，发生了严重的BUG，作者不承担任何责任。下载了该工具的用户默认已知晓上述内容。

1. Just A Simple Java Project use Maven import .jar
2. This is my simple poi-util in export .docx and you can DIY to adapt your project requirement.
3. The disclaimer asserts that the author won't be held responsible for any inaccuracies
