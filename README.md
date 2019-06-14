# Python2.7-Selenium
公司web自动化项目，基于python2.7和selenium，该版本能正常运行，包含76个报表的核对以及约20个流程；项目目录说明如下：
1、EagleProjecttest目前为上层项目控制目录，包含case、control、data、error、log、report
  case目录为测试用例py文件放置目录
  report目录下report.py为为报告生成目录，结合了unittest模块和htmltestrunner
  data为上层驱动数据目录以及比对结果数据
  error为错误截图目录
  log为运行日志目录
  report为html报告放置目录
2、EagleTry为新手练习selenium定位所写的线性流程
3、Page_object为页面封装部分
  Page_object为封装页面的类方法py
  Data为定位信息文件放置目录
4、Public_Theory文件或util为共管方法目录
  目前包含截图、删除文件、execl文件读取、获取文件路径、获取系统用户名、日志logging封装等
  此外还包含了部分后去需要实现的panda数据读取(为了比对报告)以及word文件读取
  
