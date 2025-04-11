单验报告的数据自动提取输入，单验报告的图片插入。



图片文本信息的提取使用github开源软件umi-ocr调用本地接口实现，使用教程请参考umi-ocr教程。



![image](https://github.com/user-attachments/assets/78437330-cbb0-49cc-b66c-65247283f172)# Ssv-Semi-Automatic


![image](https://github.com/user-attachments/assets/0c162031-3298-447c-9beb-66e970654f51)

dt：拉网数据截图需要自己截图保存到这个文件夹下

Pictuer：CQT截图需要自己截图保存到这个文件夹下。

jietu：是把Pictuer中的每一张信息中的、RSRP  SINR  速率信息等截取保存到jietu文件下

ping：是ping截图

result：是提取jietu中每张图片的信息导出到result.xlsx表格中，可以用"="函数直接保存到报告模板.xlsx中

报告模板.xlsx:是我这边的模板，需要放置自己的模板

可以根据我这个设置路径"D:\hongzhan"，报告中需要截取的信息，图片插入位置大小在代码中根据自己需求自己设置。


我是个水货，目前只会搞代码不会做软件，等以后会做软件再上传exe


