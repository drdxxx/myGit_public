'''
时间：2020.6.18
作者：drd
'''

import os
from random import choice
import win32com.client as win
from playsound import playsound

#播放mp3和wav文件
#source_dir：音乐文件夹或一首音乐的位置
#play_mode:播放模式（音乐文件夹），默认按顺序播放
#play_count:播放多少首音乐，默认为0(play_mode='order'表示播放全部，play_mode='random'表示播放0首)
#注意：音乐文件名不能包含中文，顺序播放时play_count必须<=（音乐文件名不包含中文且后缀名是.mp3和.wav的文件总数）
def mp3_player(source_dir,play_count=0,play_mode='order'):
    if os.path.isdir(source_dir):
        music_list = os.listdir(source_dir)
        if play_mode == 'order':
            mcount =0
            for m in music_list:
                if os.path.isfile(source_dir + '\\' + m) and m[-4:] in ('.wav','.mp3'):
                    try:
                        print('正在播放音乐:', m)
                        playsound(source_dir + '\\' + m)
                        mcount += 1
                    except:
                        print('音乐播放失败:',m)
                        continue
                    if mcount == play_count:
                        break

        elif play_mode == 'random':
            is_stop = 0
            while is_stop != play_count:
                m = choice(music_list)
                if os.path.isfile(source_dir + '\\' + m) and m[-4:] in ('.wav', '.mp3'):
                    try:
                        print('正在播放音乐:', m)
                        playsound(source_dir + '\\' + m)
                        is_stop += 1
                    except:
                        print('音乐播放失败:', m)
    else:
        if source_dir[-4:] in ('.wav', '.mp3'):
            try:
                print('正在播放音乐:', os.path.basename(source_dir))
                playsound(source_dir)
            except:
                print('音乐播放失败:', os.path.basename(source_dir))


if __name__ == '__main__':
    # 将文本转为语音并播放
    mpath = input('请输入音乐文件夹或音乐的路径:')
    list = mpath.split('\\')
    mpath = list[0]
    for i in range(1,len(list)):mpath =mpath + '\\\\' + list[i]
    text = '将要播放的音乐来源是：' + os.path.basename(mpath)
    print('音乐文件/文件夹路径:',mpath)
    sound = win.Dispatch('SAPI.SpVoice')
    sound.Speak(text)
    mp3_player(mpath,10,'random')