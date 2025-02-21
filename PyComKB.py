#本类基于pyserial库对CH9329芯片的操作进行封装
import serial
import time
import math
import random
import pyautogui
import win32api

#屏幕分辨率
width = win32api.GetSystemMetrics(0)
height = win32api.GetSystemMetrics(1)

#普通键盘按键

#CH9329里的资料不全，没有介绍清楚累加和位（就是最后一个字节）的计算方法
#普通键盘按键传输累加和位的计算方法如下
#以释放全键（KB_Up）为基准，后继数据所有字节数据的和加上KB_Up原有的累加和，就是新的累加和位的值
KB_Up = bytes([0x57,0xAB,0x00,0x02,0x08, 0x00,0x00, 0x00,0x00,0x00,0x00,0x00,0x00, 0x0C]) #12

KB_A =  bytes([0x57,0xAB,0x00,0x02,0x08, 0x00,0x00, 0x04,0x00,0x00,0x00,0x00,0x00, 0x10]) #16
KB_B =  bytes([0x57,0xAB,0x00,0x02,0x08, 0x00,0x00, 0x05,0x00,0x00,0x00,0x00,0x00, 0x11]) #17
KB_C =  bytes([0x57,0xAB,0x00,0x02,0x08, 0x00,0x00, 0x06,0x00,0x00,0x00,0x00,0x00, 0x12]) #18
KB_D =  bytes([0x57,0xAB,0x00,0x02,0x08, 0x00,0x00, 0x07,0x00,0x00,0x00,0x00,0x00, 0x13]) #19
KB_E =  bytes([0x57,0xAB,0x00,0x02,0x08, 0x00,0x00, 0x08,0x00,0x00,0x00,0x00,0x00, 0x14]) #20
KB_F =  bytes([0x57,0xAB,0x00,0x02,0x08, 0x00,0x00, 0x09,0x00,0x00,0x00,0x00,0x00, 0x15]) #21
KB_G =  bytes([0x57,0xAB,0x00,0x02,0x08, 0x00,0x00, 0x0A,0x00,0x00,0x00,0x00,0x00, 0x16]) #22
KB_H =  bytes([0x57,0xAB,0x00,0x02,0x08, 0x00,0x00, 0x0B,0x00,0x00,0x00,0x00,0x00, 0x17]) #23
KB_I =  bytes([0x57,0xAB,0x00,0x02,0x08, 0x00,0x00, 0x0C,0x00,0x00,0x00,0x00,0x00, 0x18]) #24
KB_J =  bytes([0x57,0xAB,0x00,0x02,0x08, 0x00,0x00, 0x0D,0x00,0x00,0x00,0x00,0x00, 0x19]) #25
KB_K =  bytes([0x57,0xAB,0x00,0x02,0x08, 0x00,0x00, 0x0E,0x00,0x00,0x00,0x00,0x00, 0x1A]) #26
KB_L =  bytes([0x57,0xAB,0x00,0x02,0x08, 0x00,0x00, 0x0F,0x00,0x00,0x00,0x00,0x00, 0x1B]) #27
KB_M =  bytes([0x57,0xAB,0x00,0x02,0x08, 0x00,0x00, 0x10,0x00,0x00,0x00,0x00,0x00, 0x1C]) #28
KB_N =  bytes([0x57,0xAB,0x00,0x02,0x08, 0x00,0x00, 0x11,0x00,0x00,0x00,0x00,0x00, 0x1D]) #29
KB_O =  bytes([0x57,0xAB,0x00,0x02,0x08, 0x00,0x00, 0x12,0x00,0x00,0x00,0x00,0x00, 0x1E]) #30
KB_P =  bytes([0x57,0xAB,0x00,0x02,0x08, 0x00,0x00, 0x13,0x00,0x00,0x00,0x00,0x00, 0x1F]) #31
KB_Q =  bytes([0x57,0xAB,0x00,0x02,0x08, 0x00,0x00, 0x14,0x00,0x00,0x00,0x00,0x00, 0x20]) #32
KB_R =  bytes([0x57,0xAB,0x00,0x02,0x08, 0x00,0x00, 0x15,0x00,0x00,0x00,0x00,0x00, 0x21]) #33
KB_S =  bytes([0x57,0xAB,0x00,0x02,0x08, 0x00,0x00, 0x16,0x00,0x00,0x00,0x00,0x00, 0x22]) #34
KB_T =  bytes([0x57,0xAB,0x00,0x02,0x08, 0x00,0x00, 0x17,0x00,0x00,0x00,0x00,0x00, 0x23]) #35
KB_U =  bytes([0x57,0xAB,0x00,0x02,0x08, 0x00,0x00, 0x18,0x00,0x00,0x00,0x00,0x00, 0x24]) #36
KB_V =  bytes([0x57,0xAB,0x00,0x02,0x08, 0x00,0x00, 0x19,0x00,0x00,0x00,0x00,0x00, 0x25]) #37
KB_W =  bytes([0x57,0xAB,0x00,0x02,0x08, 0x00,0x00, 0x1A,0x00,0x00,0x00,0x00,0x00, 0x26]) #38
KB_X =  bytes([0x57,0xAB,0x00,0x02,0x08, 0x00,0x00, 0x1B,0x00,0x00,0x00,0x00,0x00, 0x27]) #39
KB_Y =  bytes([0x57,0xAB,0x00,0x02,0x08, 0x00,0x00, 0x1C,0x00,0x00,0x00,0x00,0x00, 0x28]) #40
KB_Z =  bytes([0x57,0xAB,0x00,0x02,0x08, 0x00,0x00, 0x1D,0x00,0x00,0x00,0x00,0x00, 0x29]) #41

KB_LEFT_CTRL =  bytes([0x57,0xAB,0x00,0x02,0x08, 0x01,0x00, 0x00,0x00,0x00,0x00,0x00,0x00, 0x0D]) #13
KB_LEFT_SHIFT = bytes([0x57,0xAB,0x00,0x02,0x08, 0x02,0x00, 0x00,0x00,0x00,0x00,0x00,0x00, 0x0E]) #14
KB_LEFT_ALT =   bytes([0x57,0xAB,0x00,0x02,0x08, 0x04,0x00, 0x00,0x00,0x00,0x00,0x00,0x00, 0x10]) #16
KB_LEFT_WIN =   bytes([0x57,0xAB,0x00,0x02,0x08, 0x08,0x00, 0x00,0x00,0x00,0x00,0x00,0x00, 0x14]) #20

KB_RIGHT_CTRL =  bytes([0x57,0xAB,0x00,0x02,0x08, 0x10,0x00, 0x00,0x00,0x00,0x00,0x00,0x00, 0x1C]) #28
KB_RIGHT_SHIFT = bytes([0x57,0xAB,0x00,0x02,0x08, 0x20,0x00, 0x00,0x00,0x00,0x00,0x00,0x00, 0x2C]) #44
KB_RIGHT_ALT =   bytes([0x57,0xAB,0x00,0x02,0x08, 0x40,0x00, 0x00,0x00,0x00,0x00,0x00,0x00, 0x4C]) #76
KB_RIGHT_WIN =   bytes([0x57,0xAB,0x00,0x02,0x08, 0x80,0x00, 0x00,0x00,0x00,0x00,0x00,0x00, 0x8C]) #140

ALL_KBKeys = [KB_A,KB_B,KB_C,KB_D,KB_E,KB_F,KB_G,KB_H,KB_I,KB_J,KB_K,KB_L,KB_M,
              KB_N,KB_O,KB_P,KB_Q,KB_R,KB_S,KB_T,KB_U,KB_V,KB_W,KB_X,KB_Y,KB_Z,
              KB_LEFT_CTRL,KB_LEFT_SHIFT,KB_LEFT_ALT,KB_LEFT_WIN,
              KB_RIGHT_CTRL,KB_RIGHT_SHIFT,KB_RIGHT_ALT,KB_RIGHT_WIN]


#相对鼠标移动

#CH9329里的资料不全，没有介绍清楚累加和位（就是最后一个字节）的计算方法
#普通鼠标移动传输累加和位的计算方法如下
#以13为基数，鼠标数据包中后继数据的第一个字节不加入计算。鼠标向左边移动，13减移动像素点。鼠标向右移动，13加像素点。
#鼠标纵轴的移动也是如此，鼠标向上移动，13减像素点。鼠标向下移动，13加像素点。
M_Base = [0x57,0xAB,0x00,0x05,0x05, 0x01,0x00, 0x00,0x00,0x00, 0x0D]
M_set = 100
M_Base[7] = M_Base[7] + M_set
M_Base[10] = M_Base[10] + M_set
M_Move = bytes(M_Base)

M_Test1 = bytes([0x57,0xAB,0x00,0x05,0x05, 0x01,0x00,0xFD,0x00,0x00, 0x0A])
M_Test2 = bytes([0x57,0xAB,0x00,0x05,0x05, 0x01,0x00,0x00,0x05,0x00, 0x12])
M_Test3 = bytes([0x57,0xAB,0x00,0x05,0x05, 0x01,0x00,0x00,0x10,0x00, 0x1D])

#鼠标按键
M_Left_Down = bytes([0x57,0xAB,0x00,0x05,0x05, 0x01,0x01, 0x00,0x00,0x00, 0x0E])
M_Right_Down = bytes([0x57,0xAB,0x00,0x05,0x05, 0x01,0x02, 0x00,0x00,0x00, 0x0F])
M_Middle_Down = bytes([0x57,0xAB,0x00,0x05,0x05, 0x01,0x04, 0x00,0x00,0x00, 0x11])
M_Button_Up = bytes([0x57,0xAB,0x00,0x05,0x05, 0x01,0x00, 0x00,0x00,0x00, 0x0D])

#绝对鼠标移动
XL = 1920
YL = 1080

M_ATest1 = bytes([0x57,0xAB,0x00,0x04,0x07, 0x02,0x00,0x40,0x01,0x15,0x02,0x00, 0x67])
M_ATest2 = bytes([0x57,0xAB,0x00,0x04,0x07, 0x02,0x00,0x19,0x0C,0x6B,0x0A,0x00, 0xA9])


#多媒体键盘按键
MM_Mute = bytes([0x57,0xAB,0x00,0x03,0x04,0x02,0x04,0x00,0x00,0x0F])
MM_Up = bytes([0x57,0xAB,0x00,0x03,0x04,0x02,0x00,0x00,0x00,0x0B])

#本类对CH9329的操作方式进行封装，方便键盘鼠标模拟
class SimuComKB():
    def __init__(self,ComName) -> str:  #构造函数
        try:
            self.__KBCom = serial.Serial(ComName,9600,timeout=3)
        except Exception as e:
            print("Error:",e)
            return

    def __del__(self):  #析构函数
        self.__KBCom.close()
        #self.KB_Up()    #释放所有键
        #self.M_BUp()

        try:        #释放COM端口
            self.__KBCom.close()
        except Exception as e:
            print("Error:",e)

    def KB_Down(self,Key:str):   #键盘按键按下
        Key = Key.lower()
        if(Key == 'a'):
            self.__KBCom.write(KB_A)
            return

        if(Key == 'b'):
            self.__KBCom.write(KB_B)
            return

        if(Key == 'c'):
            self.__KBCom.write(KB_C)
            return

        if(Key == 'd'):
            self.__KBCom.write(KB_D)
            return

        if(Key == 'e'):
            self.__KBCom.write(KB_E)
            return

        if(Key == 'f'):
            self.__KBCom.write(KB_F)
            return

        if(Key == 'g'):
            self.__KBCom.write(KB_G)
            return

        if(Key == 'h'):
            self.__KBCom.write(KB_H)
            return

        if(Key == 'i'):
            self.__KBCom.write(KB_I)
            return

        if(Key == 'j'):
            self.__KBCom.write(KB_J)
            return

        if(Key == 'k'):
            self.__KBCom.write(KB_K)
            return

        if(Key == 'l'):
            self.__KBCom.write(KB_L)
            return

        if(Key == 'm'):
            self.__KBCom.write(KB_M)
            return

        if(Key == 'n'):
            self.__KBCom.write(KB_N)
            return

        if(Key == 'o'):
            self.__KBCom.write(KB_O)
            return

        if(Key == 'p'):
            self.__KBCom.write(KB_P)
            return

        if(Key == 'q'):
            self.__KBCom.write(KB_Q)
            return

        if(Key == 'r'):
            self.__KBCom.write(KB_R)
            return

        if(Key == 's'):
            self.__KBCom.write(KB_S)
            return

        if(Key == 't'):
            self.__KBCom.write(KB_T)
            return

        if(Key == 'u'):
            self.__KBCom.write(KB_U)
            return

        if(Key == 'v'):
            self.__KBCom.write(KB_V)
            return

        if(Key == 'w'):
            self.__KBCom.write(KB_W)
            return

        if(Key == 'x'):
            self.__KBCom.write(KB_X)
            return

        if(Key == 'y'):
            self.__KBCom.write(KB_Y)
            return

        if(Key == 'z'):
            self.__KBCom.write(KB_Z)
            return

        if(Key == 'left_shift'):
            self.__KBCom.write(KB_LEFT_SHIFT)
            return
        
        if(Key == 'left_win'):
            self.__KBCom.write(KB_LEFT_WIN)
            return

        if(Key == 'left_ctrl'):
            self.__KBCom.write(KB_LEFT_CTRL)
            return

        if(Key == 'left_alt'):
            self.__KBCom.write(KB_LEFT_ALT)
            return

        if(Key == 'right_shift'):
            self.__KBCom.write(KB_RIGHT_SHIFT)
            return
        
        if(Key == 'right_win'):
            self.__KBCom.write(KB_RIGHT_WIN)
            return

        if(Key == 'right_ctrl'):
            self.__KBCom.write(KB_RIGHT_CTRL)
            return

        if(Key == 'right_alt'):
            self.__KBCom.write(KB_RIGHT_ALT)
            return

    def KB_TestAll(self):   #测试所有键盘按键
        for key in ALL_KBKeys:
            self.__KBCom.write(key)
            self.__KBCom.write(KB_Up)

    def KB_Up(self):    #键盘按键抬起
        self.__KBCom.write(KB_Up)

    def KB_Click(self,Key:str):     #键盘按键点击
        self.KB_Down(Key)
        time.sleep(0.1)
        self.KB_Up()

    def KB_ClickDely(self,Key:str,Dely):        #按下Dely秒键盘按键
        self.KB_Down(Key)
        time.sleep(Dely)
        self.KB_Up()

    def scroll(self, y):
        if y >= 0:
            y_cur = y
        elif y < 0:
            y_cur = 0xFF + y
        footer = (269 + y_cur) % 256
        send_Base = [0x57, 0xAB, 0x00, 0x05, 0x05, 0x01, 0x00, 0x00, 0x00, y_cur, footer]
        self.__KBCom.write(bytes(send_Base))

    def move(self, XL: int, YL: int):
        if XL >= 0:
            X_cur = XL
        elif XL < 0:
            X_cur = 0xFF + XL
        if YL >= 0:
            Y_cur = YL
        elif YL < 0:
            Y_cur = 0xFF + YL
        footer = (269 + X_cur + Y_cur) % 256
        send_Base = [0x57, 0xAB, 0x00, 0x05, 0x05, 0x01, 0x00, X_cur, Y_cur, 0x00, footer]
        self.__KBCom.write(bytes(send_Base))

    def M_MoveR(self,x:int,y:int):  #鼠标相对移动

        abs_x = abs(x)
        abs_y = abs(y)
        x_cnt = abs_x // 10
        x_last = abs_x % 10
        y_cnt = abs_y // 10
        y_last = abs_y % 10
        if x_cnt >= y_cnt:
            temp = x_cnt - y_cnt
            while y_cnt > 0:
                self.move(10 * x // abs_x, 10 * y // abs_y)
                y_cnt -= 1
            while temp > 0:
                self.move(10 * x // abs_x, 0)
                temp -= 1
        elif x_cnt < y_cnt:
            temp = y_cnt - x_cnt  # 修正此处逻辑
            while x_cnt > 0:
                self.move(10 * x // abs_x, 10 * y // abs_y)
                x_cnt -= 1
            while temp > 0:
                self.move(0, 10 * y // abs_y)
                temp -= 1
        if x_last != 0:
            self.move(x_last * x // abs_x, 0)
        if y_last != 0:
            self.move(0, y_last * y // abs_y)

    def M_RandMove(self):       #鼠标随机移动
        self.M_MoveR(random.randint(-500,700),random.randint(-400,600))

    #pX是横坐标，pY是纵坐标
    def M_MoveTo(self,pX:int,pY:int):     #鼠标绝对移动
        X_Cur = int((4096 * pX) / width)
        Y_Cur = int((4096 * pY) / height)
        X_high = (X_Cur >> 8) & 0xFF
        X_low = X_Cur & 0xFF
        Y_high = (Y_Cur >> 8) & 0xFF
        Y_low = Y_Cur & 0xFF
        footer = (271 + X_high + X_low + Y_high + Y_low) % 256
        send_Base = [0x57, 0xAB, 0x00, 0x04, 0x07, 0x02, 0x00, X_low, X_high, Y_low, Y_high, 0x00, footer]
        self.__KBCom.write(bytes(send_Base))

    def M_MoveToBox(self,points:list,Cell=40):      #鼠标移动到一个指定的区域内
        if(len(points) != 4):
            print("The Param [points] must have 4 numbers!")
            return

        #计算区域中点
        pX = points[0][0] + points[1][0]+ points[2][0] + points[3][0]
        pY = points[0][1] + points[1][1]+ points[2][1] + points[3][1]
        pX = math.ceil(pX/4)
        pY = math.ceil(pY/4)

        #移动鼠标
        self.M_MoveTo(pX,pY,Cell)

    def M_MoveToRandBox(self,points:list,Cell=40):      #鼠标移动到一个指定区域中的任意一点
        if(len(points) != 4):
            print("The Param [points] must have 4 numbers!")
            return

        #计算区域中点
        pX = points[0][0] + points[1][0]+ points[2][0] + points[3][0]
        pY = points[0][1] + points[1][1]+ points[2][1] + points[3][1]
        pX = pX/4
        pY = pY/4

        #在区域内偏移中点值
        pX = (pX + random.randint(points[0][0],points[1][0]))/2
        pX = math.ceil(pX)
        pY = (pY + random.randint(points[0][1],points[3][1]))/2
        pY = math.ceil(pY)

        #移动鼠标
        self.M_MoveTo(pX,pY,Cell)

    def M_BDown(self,Key:str):      #鼠标按钮按下
        Key = Key.lower()
        if(Key == 'l'):
            self.__KBCom.write(M_Left_Down)
        if(Key == 'r'):
            self.__KBCom.write(M_Right_Down)
        if(Key == 'm'):
            self.__KBCom.write(M_Middle_Down)

        pass
       
    def M_BUp(self):            #鼠标按钮抬起
        self.__KBCom.write(M_Button_Up)     

    def M_BClick(self,Key:str):     #鼠标按钮点击
        self.M_BDown(Key)
        time.sleep(0.1)
        self.M_BUp()

    def M_BDClick(self,Key:str):        #鼠标双击
        self.M_BDown(Key)
        self.M_BUp()
        self.M_BDown(Key)
        self.M_BUp()
        
    


