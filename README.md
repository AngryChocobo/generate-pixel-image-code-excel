# 拼豆像素图对应的excel表生成器

## 背景
哥们开始玩拼豆了，想把一些简单的像素图拼好，但是找到的图片需要肉眼与拼豆颜色进行对应，非常累
![4257cfaaa548628b7f69755819e56971](https://github.com/user-attachments/assets/1e5880c8-a41d-488d-bcd7-4c126c0913a9)

![62c105290800fd4657bd9f088dfc9ed3](https://github.com/user-attachments/assets/2d95922e-c041-463f-ab53-b495ea8e3d2e)

## 程序逻辑

有一些部分我们可以优化，比如：
1. 分析图片的像素点，找到与颜色对应的拼豆的颜色和编码
2. 生成类似excel的结构，方便对照编码来拼图



## 实现流程
1. 商家没提供拼豆对应的rgb值，所以我将比较规则的拼豆宣传图按照每个`碗`进行切割，并计算`碗`内像素的平均值，生成`用户颜色list`
2. 宣传图没有对应编码，所以根据实物图+OCR形式获取拼豆编码
3. 写程序遍历文件夹内所有图片，提取出去重后的像素map，并使用多种算法取的该像素点在`用户颜色list`中最接近的颜色，绘制到单元格的背景色，同时单元格记录颜色编码，每种算法单独按sheet页绘制
4. 交给哥们 对着excel，扣掉边边角角不需要的区域，手动拼图

## 运行方式

1. 自行安装python虚拟运行环境
2. pip install -r requirements.txt
3. python main.py （读取imgs内的所有图片）


## 效果图


![03433d5b21bd81f44d13f47e83e44de6](https://github.com/user-attachments/assets/613fff6d-5212-4e28-95e1-661cab2cad97)

![0d277a0d4a199bd85a484fe8808350d4](https://github.com/user-attachments/assets/0949038c-20a7-4177-84f1-4d844e570421)
