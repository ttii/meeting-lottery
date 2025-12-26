#!/usr/bin/python
# -*- coding: utf8 -*-
"""
powerby toby 2025-12-23
吉客道年会抽奖程序,通过pygame渲染图片获得点击事件,随机抽奖
1. esc按钮退出程序后,会生成"中奖名单.xls"文件
2. F5按钮可以最小化抽奖程序,并停止播放声音
"""
import secrets
import sys
import xlrd
import pygame
import random

import xlwt
from pygame.locals import *

# fullScreen = True
mode = 0
mouse_scan = [(1451, 498, 300, 300), (1451, 498, 300, 300)]
pos_scan = [(84, 814, 960, 886), (1513, 65, 914, 760)]
top_center_pos = (950, 157, 960, 886)
left_center_pos = (163, 367, 960, 886)
right_left_pos = (1450, 6, 914, 760)
audio_enabled = False


def get_name_list_from_excel(file_name):
    """解析 人员.xlsx 文件，得到人员名单列表"""
    name_list = []
    excelFile = xlrd.open_workbook(file_name)
    sheet = excelFile.sheet_by_name('Sheet1')
    print(sheet.name, sheet.nrows, sheet.ncols)
    job_num = sheet.cell(0, 0).value
    job_name = sheet.cell(0, 1).value
    for row in range(1, sheet.nrows):
        job_num = sheet.cell(row, 0).value
        job_name = sheet.cell(row, 1).value
        # print(job_num, job_name)
        name_list.append((job_num, job_name))
    return job_num, job_name, name_list


def handle_mouse_event(index, pause_flag):
    if pygame.mouse.get_pressed()[0]:
        x, y = pygame.mouse.get_pos()
        x -= m.get_width() / 2
        y -= m.get_height() / 2
        if (mouse_scan[mode][0] + mouse_scan[mode][2] > x > mouse_scan[mode][0] and
            mouse_scan[mode][1] + mouse_scan[mode][3] > y > mouse_scan[mode][1]):
            pause_flag = not pause_flag
            global audio_enabled
            if not pause_flag:
                if audio_enabled:
                    pygame.mixer.music.play()
                winner = name_list[index]
                winner_name_list.append(winner)
                print(f'本轮中奖者:{winner}')
                del name_list[index]
            else:
                if audio_enabled:
                    pygame.mixer.music.stop()

    return index, pause_flag


def show_name_list(name_list):
    """调试接口，输出当前名单"""
    for index in range(0, len(name_list)):
        s = "{} {}".format(name_list[index][0], name_list[index][1])
        print(s)

def _show_info_text(text, font_size, font_color, background_color, pos_x, pos_y):
    font = pygame.font.Font("simhei.ttf", font_size)
    text_obj = font.render(text, True, font_color, background_color)
    text_pos = text_obj.get_rect()
    text_pos.center = (pos_x, pos_y)
    screen.blit(text_obj, text_pos)

def _show_info_text_left(text, font_size, font_color, background_color, pos_x, pos_y):
    font = pygame.font.Font("simhei.ttf", font_size)
    text_obj = font.render(text, True, font_color, background_color)
    text_pos = text_obj.get_rect()
    text_pos.topleft = (pos_x, pos_y)
    screen.blit(text_obj, text_pos)


def save_winners_to_excel(winner_list, filename = '中奖名单.xls'):
    """保存中奖者名单到Excel文件"""
    workbook = xlwt.Workbook(encoding = 'utf-8')
    worksheet = workbook.add_sheet('中奖名单')

    # 设置表头
    headers = ['工号', '姓名']
    for col, header in enumerate(headers):
        worksheet.write(0, col, header)

    # 写入中奖者数据
    for row, (job_num, job_name) in enumerate(winner_list, start = 1):
        worksheet.write(row, 0, job_num)
        worksheet.write(row, 1, job_name)

    workbook.save(filename)
    print(f'中奖者名单已保存到 {filename}')


if __name__ == "__main__":
    job_num, job_name, name_list = get_name_list_from_excel(r'name_file.xls')
    print(len(name_list))
    winner_name_list = []


    pygame.init()
    mg = 'gc_cz.png'

    # if fullScreen:
    mode = 0
    bg = 'bg_1920x1080.jpg'
    screen = pygame.display.set_mode((1920, 1080), FULLSCREEN, 32)

    pygame.display.set_caption("吉客道年会抽奖程序 Powerby Toby")

    b = pygame.image.load(bg).convert()
    m = pygame.image.load(mg).convert_alpha()
    screen.blit(b, (0, 0))
    screen.blit(m, (0, 0))

    # 将音频相关代码注释掉或用条件判断跳过
    try:
        pygame.mixer.init()
        pygame.mixer.music.load('run.wav')
        audio_enabled = True
    except pygame.error:
        print("音频初始化失败，禁用音频功能")

    # 在需要播放音频的地方：
    if audio_enabled:
        pygame.mixer.music.play()

    # index = random.randint(0, len(name_list) - 1)
    index = secrets.randbelow(len(name_list))
    pause_flag = False
    running = True
    enough = True
    while True:
        if pygame.display.get_active():
            running = True
        for event in pygame.event.get():
            if event.type == pygame.KEYDOWN and event.key == pygame.K_ESCAPE:
                pygame.quit()
                if winner_name_list:
                    save_winners_to_excel(winner_name_list)
                sys.exit(0)
            elif event.type == pygame.KEYDOWN and event.key == pygame.K_F5:
                if audio_enabled:
                    running = False
                    pygame.mixer.music.stop()

                pygame.display.iconify()
            elif event.type == MOUSEBUTTONDOWN:
                if enough:
                    index, pause_flag = handle_mouse_event(index, pause_flag)

        screen.blit(b, (0, 0))
        x, y = pygame.mouse.get_pos()
        x -= m.get_width() / 2
        y -= m.get_height() / 2
        pygame.mouse.set_visible(False)
        screen.blit(m, (x, y))

        if not running:
            continue

        if not pause_flag:
            if enough:
                if len(name_list) < 2:
                    enough = False
                    continue

                index = random.randint(0, len(name_list)-1)

                if audio_enabled and not pygame.mixer.music.get_busy():
                    pygame.mixer.music.play()

        if enough:
            text_context = '{} {}'.format(name_list[index][0], name_list[index][1])
            # print(text_context)
            _show_info_text(text_context, 80, (205, 0, 0), (253, 244, 180), pos_x=top_center_pos[0], pos_y=top_center_pos[1])
        else:
            text_context = '参与人数不足2人,抽奖结束!'
            # print(text_context)
            _show_info_text(text_context, 80, (205, 0, 0), (253, 244, 180), pos_x = top_center_pos[0], pos_y = top_center_pos[1])

        _show_info_text(f'当前参与人数: {len(name_list)}', 30, (253, 244, 180), (205, 0, 0), pos_x=left_center_pos[0], pos_y=left_center_pos[1])


        _show_info_text_left(f'历史中奖者:', 30, (253, 244, 180), (205, 0, 0), pos_x=right_left_pos[0], pos_y=right_left_pos[1])

        for i in range(0, len(winner_name_list)):
            _pos = right_left_pos[1]
            _pos += (i+1)*32
            _info = winner_name_list[i]
            _show_info_text_left(f'{_info[0]} {_info[1]}', 30, (253, 244, 180), (205, 0, 0), pos_x=right_left_pos[0]+60, pos_y=_pos)


        pygame.display.update()
