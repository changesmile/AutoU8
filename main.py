# coding=utf-8
import json
import datetime
import os
import sys
import pyautogui
import time
import requests

# width, height = pyautogui.size()  # 获取屏幕的高和宽
# print(width, height)
date01 = datetime.date.today()
# 将此处的机器人hook地址替换为你创建的机器人地址即可
webhook_url = "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=c7cec630-bc87-4c2a-9e1b-bef8c1194fa8"
text_push_content = """密码过期"""

text_data = {
    "msgtype": "text",
    "text": {
        "content": text_push_content,
        "mentioned_list": ["@all"]
    }
}


def click_win(name, click_xy):
    if name == "login":
        time.sleep(2)
        pyautogui.click(click_xy)
        pyautogui.click(click_xy)
        time.sleep(10)
        xy = pyautogui.locateOnScreen(f"D:/image/password.png")  # 左边的x坐标、顶边的y坐标、宽度以及高度
        if xy is not None:
            PostData = str(json.dumps(text_data)).encode('utf-8')
            r = requests.post(webhook_url, data=PostData)

    elif name == "out_file" or name == "out_file2" or name == "out_file3":

        pyautogui.click(click_xy)
        time.sleep(60)
        pyautogui.press('right')  # 按下左
        time.sleep(1)

        time.sleep(1)
        pyautogui.typewrite(message=str(date01))  # 输入当前日期
        pyautogui.hotkey(' ')
        pyautogui.hotkey('Enter')  # 配套使用确定保存文件
        time.sleep(1)
        pyautogui.press("n")  # 键盘输入空格
        time.sleep(1)
        pyautogui.hotkey('esc')
        time.sleep(5)
        pyautogui.hotkey(' ')
        pyautogui.hotkey('Enter')  # 确定文件输出成功的弹窗
    elif name == "out_file5":
        pyautogui.click(click_xy)
        time.sleep(60)
        pyautogui.press('left')  # 按下左
        time.sleep(30)
        pyautogui.hotkey(' ')
        pyautogui.hotkey('Enter')  # 配套使用确定保存文件
        time.sleep(30)
        pyautogui.hotkey(' ')
        pyautogui.hotkey('Enter')  # 配套使用确定保存文件
        time.sleep(3)
        pyautogui.hotkey('esc')
        time.sleep(5)

    elif name == "confirm":
        pyautogui.click(click_xy)
        time.sleep(40)

    elif name == "out_file5" or name == "confirm3":
        pyautogui.click(click_xy)
        time.sleep(3)
    else:
        pyautogui.click(click_xy)
        time.sleep(15)


def AutoU8(list_U8):
    pyautogui.hotkey(' ')
    pyautogui.hotkey('Enter')  # 配套使用确定保存文件
    os.system(r"start D:\U8SOFT\EnterprisePortal.exe")  # 自启默认在D盘下的u8
    time.sleep(3)
    j = 0
    for list_eve in list_U8:
        for i in list_eve:

            while j < 3:  # 设置三次检测图像
                xy = pyautogui.locateOnScreen(f"D:/image/{i}.png")  # 左边的x坐标、顶边的y坐标、宽度以及高度
                if xy is not None:
                    click_win(i, xy)
                    break
                if xy is None:
                    time.sleep(60)  # 检测失败进入休眠
                    print("失败进来了")
                j = j + 1
            if j == 3:
                return True  # 三次失败则结束当前程序
            j = 0


if __name__ == '__main__':
    # 删除三个请购和采购表格，不然U8导入时候覆盖会出问题
    try:
        os.remove('E:\ERP Data\请购执行进度表.xlsx')
    except Exception as e:
        print(e)
    try:
        os.remove('E:\ERP Data\请购单列表.xlsx')
    except Exception as e:
        print(e)
    try:
        os.remove('E:\ERP Data\采购订单列表.xlsx')
    except Exception as e:
        print(e)

    for i in range(5):
        # 三个list,新增close窗口，回车直接确定
        # 三个list结束才能kill

        # 每次间隔时间设置
        list_all = []
        Start_list = ['login']
        MRP_list = ['mrp', 'search', 'confirm', 'out_file', 'close']  # MRP维护计划
        Inventory_list = ['Inventory', 'search2', 'confirm', 'out_file2', 'close']  # 存货档案
        Dispatch_list = ['Dispatch', 'search3', 'confirm', 'out_file3', 'close']  # 工序派工
        PurchaseOrder_list = ['PurchaseOrder', 'search4', 'confirm', 'out_file5', 'close']  # 采购订单列表
        PurchaseRequisition_list = ['PurchaseRequisition', 'search4', 'confirm', 'out_file5', 'close']  # 请购单列表
        PurchaseSchedule_list = ['PurchaseSchedule', 'confirm', 'out_file6', 'confirm3', 'out_file7']  # 请购执行进度表
        list_all.append(Start_list)
        if not os.path.isfile(f'E:\ERP Data\MRP计划维护--全部{date01}.XLSX'):
            list_all.append(MRP_list)
        if not os.path.isfile(f'E:\ERP Data\存货档案{date01}.XLSX'):
            list_all.append(Inventory_list)
        if not os.path.isfile(f'E:\ERP Data\工序派工资料维护{date01}.XLSX'):
            list_all.append(Dispatch_list)
        list_all.append(PurchaseOrder_list)
        list_all.append(PurchaseRequisition_list)
        list_all.append(PurchaseSchedule_list)

        for i in range(len(list_all)):
            if AutoU8(list_all):
                os.system("taskkill /F /IM EnterprisePortal.exe")  # 杀死U8进程
                AutoU8(list_all)
            else:
                break
        list_all.clear()
        time.sleep(120)
        os.system("taskkill /F /IM EnterprisePortal.exe")  # 杀死U8进程
        sys.exit()  # 结束程序
