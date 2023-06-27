# -*- coding: utf-8 -*-
###########################################################################
#    Copyright 2023 奇客罗方智能科技 https://www.geekercloud.com
#    ITGeeker.net <alanljj@gmail.com>
############################################################################
import base64
import json
import os
import shutil
import tempfile
import tkinter as tk
from tkinter import ttk
# import ttkbootstrap as ttk
from tkinter import messagebox
import tkinter.filedialog
from webbrowser import open_new_tab

from str_replacement_api import generate_file_and_str_list


class ListTreeSheet(tk.Frame):

    def __init__(self, master):
        super().__init__(master)
        # self.label1 = tk.Label(self, text='待替换文字列表')
        # self.label1.config(font=('Microsoft YaHei UI', 14))
        # self.label1.grid(row=0, column=0, columnspan=2, ipadx=10, ipady=10)

        self.org_str = None
        self.replaced_str = None
        self.start_remove_button = None
        self.browse_button = None
        self.edit_button = None
        self.listbox = None
        self.entry_path = None
        self.add_button = None

        self.list_frame()
        self.string_frame()
        self.select_path_frame()
        self.author_frame()

    def add_item(self):
        org_str = self.org_str.get()
        replaced_str = self.replaced_str.get()
        if not org_str:
            messagebox.showwarning(title="温馨提醒", message="请输入要查找的原始文字！")
        if org_str in replaced_str:
            messagebox.showwarning(title="温馨提醒", message="替换内容不能包含原始内容！")
        else:
            row_values = [org_str, replaced_str]
            # Insert row into treeview
            self.treeview.insert('', tk.END, values=row_values)
            self.org_str.delete(0, "end")
            self.replaced_str.delete(0, "end")

    def edit_item(self):
        selected_value_l = []
        # selected_value_l = self.treeview.focus()
        selected_items = self.treeview.selection()
        if selected_items:
            self.org_str.delete(0, "end")
            self.replaced_str.delete(0, "end")
            for sitem in selected_items:
                svalue = self.treeview.item(sitem)
                print('svalue: %s' % type(svalue))  # <class 'dict'>
                print('svalue: %s' % svalue)
                # print('svalue.details: %s' % svalue.get('values'))
                selected_value_l = svalue.get('values')
                print('selected_value_l: %s' % selected_value_l)
                if len(selected_value_l) == 1:
                    self.org_str.insert(0, selected_value_l[0])
                elif len(selected_value_l) == 2:
                    self.org_str.insert(0, selected_value_l[0])
                    self.replaced_str.insert(0, selected_value_l[1])
                self.treeview.delete(sitem)
        else:
            messagebox.showwarning(title="温馨提醒", message="请先选择想要修改的文字内容！")

    def get_all_tree_view_list(self):
        val_list_all = []
        for child in self.treeview.get_children():
            item_dict = dict()
            # print(self.treeview.item(child)["values"])
            # val_list_all.append(self.treeview.item(child)["values"])
            child_val_l = self.treeview.item(child)["values"]
            if len(child_val_l) == 1:
                item_dict.update({'org_str': child_val_l[0]})
            elif len(child_val_l) == 2:
                item_dict.update({
                    'org_str': child_val_l[0],
                    'replaced_str': child_val_l[1]
                })
            val_list_all.append(item_dict)
        print('val_list_all: %s' % val_list_all)
        return val_list_all

    # def open_popup():
    #     top = Toplevel(win)
    #     top.geometry("750x250")
    #     top.title("Child Window")
    #     Label(top, text="Hello World!", font=('Mistral 18 bold')).place(x=150, y=80)
    #
    # Label(win, text=" Click the Below Button to Open the Popup Window", font=('Helvetica 14 bold')).pack(pady=20)
    # # Create a button in the main Window to open the popup
    # ttk.Button(win, text="Open", command=open_popup).pack()

    def start_replace_strings_from_path(self):
        val_list = self.get_all_tree_view_list()
        if val_list:
            self.save_all_item_to_json(val_list)
        if not val_list:
            messagebox.showwarning(title="温馨提醒", message="请添加要替换的文字！")
        elif not self.entry_path.get() or self.entry_path.get() == '浏览并选择目录':
            messagebox.showwarning(title="温馨提醒", message="请先选择文件的目录！")
        else:
            print('self.entry_path.get(): %s' % self.entry_path.get())
            print('val_list: %s' % val_list)

            finished_nmb = generate_file_and_str_list(self.entry_path.get(), val_list)
            if finished_nmb:
                messagebox.showinfo(title="任务通知", message="任务已圆满完成！共处理了%s个文件,\n"
                                                              "已替换的文件带有-revised字样，保存在子目录《已替换文件》中。"
                                                              % str(finished_nmb))
            else:
                messagebox.showerror(title="任务错误通知", message="任务完成，但有错误！")

    def generate_json_ffp(self):
        cur_usr_path = os.environ['USERPROFILE']
        # print('cur_usr_path: %s' % cur_usr_path)
        replace_str_f = os.path.join(cur_usr_path, 'itgeeker_word_str_replacement.json')
        if not os.path.isfile(replace_str_f):
            ffp_d = dict()
            with open(replace_str_f, 'w', encoding='utf-8') as fp:
                fp.write(json.dumps(ffp_d, indent=4, ensure_ascii=False))
                # pass
        return replace_str_f

    def save_all_item_to_json(self, value_list):
        print("here should to save all")
        ffp_d = dict()
        replace_str_f = self.generate_json_ffp()

        if self.entry_path.get():
            ffp_d['entry_path'] = self.entry_path.get()

        # print('ffp_d: ', ffp_d)
        with open(replace_str_f, 'w', encoding='utf-8') as ffp:
            ffp_d['string_list'] = value_list
            ffp.write(json.dumps(ffp_d, indent=4, ensure_ascii=False))

    def read_all_item_to_treeview_list(self):
        json_ffp = self.generate_json_ffp()
        if json_ffp:
            with open(json_ffp, 'r', encoding='utf-8') as ffp:
                dt_dict = json.load(ffp)
                print('dt_dict: %s' % dt_dict)
            # # keys = tuple(dt.keys())
            # keys = ('原始文字', '替换文字')
            # for col_name in keys:
            #     self.treeview.heading(col_name, text=col_name)
            if 'string_list' in dt_dict:
                self.treeview.delete(*self.treeview.get_children())
                for dt in dt_dict['string_list']:
                    value_tuple = tuple(dt.values())
                    self.treeview.insert('', tk.END, values=value_tuple)
            if 'entry_path' in dt_dict:
                self.entry_path.delete(0, tk.END)
                self.entry_path.insert(0, dt_dict['entry_path'])
                self.folder_info_fram()
                self.cout_nmb_of_doc(dt_dict['entry_path'])
            if 'include_sub_dir' in dt_dict:
                print('include_sub_dir: ', dt_dict['include_sub_dir'])
                self.include_sub_dir.set(dt_dict['include_sub_dir'])
            if 'label_file_nmb' in dt_dict:
                # print('label_file_nmb: ', dt_dict['label_file_nmb'])
                if dt_dict['label_file_nmb']:
                    self.label_file_nmb.config(text='文件数：' + str(dt_dict['label_file_nmb']))

    def delete_items(self, _):
        print('delete')
        for i in self.treeview.selection():
            self.treeview.delete(i)

    def cout_nmb_of_doc(self, directory):
        count_docx = 0
        count_doc = 0
        for root_dir, cur_dir, files in os.walk(directory):
            for file in files:
                # ffp = os.path.join(root_dir, file)
                if file.endswith(".docx"):
                    count_docx += 1
                elif file.endswith(".doc"):
                    count_doc += 1
        print('file count_docx:', count_docx)
        print('file count_doc:', count_doc)
        # return count_docx, count_doc
        self.label_docx_nmb.config(text='.docx文件：' + str(count_docx) + '个')
        self.label_doc_nmb.config(text='.doc文件：' + str(count_doc) + '个')
        if count_doc:
            self.label_doc_nmb.config(text='.doc文件：' + str(count_doc) + '个 (本工具只支持处理docx格式文件)')
            label_doc_reminder = ttk.Label(self.folder_frame,
                                           text='建议使用技术奇客的开源工具 - Word格式转换(.doc ➡ .docx) [点击下载并安装]',
                                           cursor="hand2")
            label_doc_reminder.config(font=('Microsoft YaHei UI', 10))
            label_doc_reminder.bind("<Button-1>",
                                    lambda e: self.open_website(
                                        "https://www.itgeeker.net/itgeeker-technical-service"
                                        "/itgeeker_convert_doc_to_docx/"))
            label_doc_reminder.grid(row=1, column=0, columnspan=2, padx=(15, 10), ipadx=10, ipady=5, sticky="ew")

    def select_directory(self):
        directory = tk.filedialog.askdirectory()
        self.entry_path.delete(0, tk.END)
        self.entry_path.insert(0, directory)
        self.folder_info_fram()
        self.cout_nmb_of_doc(directory)

    def open_website(self, url):
        open_new_tab(url)

    def on_window_close(self):
        print("Window closed")
        val_list_on_close = self.get_all_tree_view_list()
        self.save_all_item_to_json(val_list_on_close)
        geekerWin.destroy()

    def list_frame(self):
        list_replace_frame = ttk.LabelFrame(self, text="文字内容列表")
        list_replace_frame.grid(row=0, column=0, columnspan=2, sticky='nsew')

        # list_files_frame = ttk.LabelFrame(self, text="Word旧格式.doc文件, 可多选")
        list_files_frame = ttk.Frame(list_replace_frame)
        list_files_frame.grid(row=0, column=0, columnspan=3, padx=10, pady=10, ipadx=10, sticky='nsew')

        cols = ("原始内容", "替换内容")
        self.treeview = ttk.Treeview(list_files_frame, show="headings", columns=cols, height=9, selectmode="browse")
        self.treeview.column("# 1", anchor="w", width=428)
        self.treeview.heading("# 1", text="原始内容", anchor="w")
        self.treeview.column("# 2", anchor="w", width=288)
        self.treeview.heading("# 2", text="替换内容", anchor="w")
        # self.treeview.bind("<Return>", lambda e: self.select_multiple())
        # self.treeview.bind('<ButtonRelease-1>', self.cur_select_item)
        # self.treeview.bind('<ButtonRelease-1>', self.select_multiple)
        self.treeview.bind('<Delete>', self.delete_items)
        self.treeview.pack(expand=True, fill='both')

        tree_y_scroll = ttk.Scrollbar(self.treeview, orient='vertical', command=self.treeview.yview)
        self.treeview.configure(yscrollcommand=tree_y_scroll.set)
        tree_y_scroll.place(relx=1, rely=0, relheight=1, anchor='ne')
        # mousewheel scrolling
        self.treeview.bind('<MouseWheel>', lambda event: self.treeview.yview_scroll(-int(event.delta / 60), "units"))

        tree_x_scroll = ttk.Scrollbar(self.treeview, orient='horizontal', command=self.treeview.xview)
        self.treeview.configure(xscrollcommand=tree_x_scroll.set)
        tree_x_scroll.place(relx=0, rely=1, relwidth=1, anchor='sw')
        # event to scroll left / right on Ctrl + mousewheel
        self.treeview.bind('<Control MouseWheel>',
                           lambda event: self.treeview.xview_scroll(-int(event.delta / 60), "units"))

    def string_frame(self):
        string_frame = ttk.LabelFrame(self, text="文字调整")
        string_frame.grid(row=2, column=0, columnspan=2, padx=10, pady=5, ipadx=10, sticky='nsew')

        str_org_frame = ttk.LabelFrame(string_frame, text="原始文字")
        str_org_frame.grid(row=0, column=0, padx=10, pady=5, ipadx=10, sticky='nsew')

        self.org_str = tk.Entry(str_org_frame, justify=tk.CENTER, width=32,
                                  font=('Microsoft YaHei UI', 12))
        # self.org_str.insert(0, "输入您想查找的字符")
        self.org_str.focus_force()
        self.org_str.grid(row=0, column=0, padx=10, pady=5, sticky='nsew')

        str_replaced_frame = ttk.LabelFrame(string_frame, text="替换文字")
        str_replaced_frame.grid(row=0, column=1, padx=10, pady=5, ipadx=10, sticky='nsew')

        self.replaced_str = tk.Entry(str_replaced_frame, justify=tk.CENTER, width=32,
                                  font=('Microsoft YaHei UI', 12))
        # self.replaced_str.insert(0, "输入您想替换的字符")
        # self.replaced_str.focus_force()
        self.replaced_str.grid(row=0, column=0, padx=10, pady=5, sticky='nsew')

        self.add_button = tk.Button(string_frame, text="添加替换", command=self.add_item, bg='brown', fg='white',
                                    font=('Microsoft YaHei UI', 11, 'bold'))
        self.add_button.grid(row=1, column=0, padx=10, pady=5, ipadx=10)

        self.edit_button = tk.Button(string_frame, text="编辑或删除", command=self.edit_item, bg='grey', fg='white',
                                     font=('Microsoft YaHei UI', 11, 'normal'))
        self.edit_button.grid(row=1, column=1, padx=10, pady=5, ipadx=10)

    def folder_info_fram(self):
        self.folder_frame = ttk.LabelFrame(self, text="目录信息")
        self.folder_frame.grid(row=3, column=0, columnspan=2, padx=10, pady=10, ipadx=10, sticky='nsew')

        self.label_docx_nmb = ttk.Label(self.folder_frame, text='.docx文件数')
        self.label_docx_nmb.config(font=('Microsoft YaHei UI', 10))
        self.label_docx_nmb.configure(justify="center", anchor="e")
        self.label_docx_nmb.grid(row=0, column=0, padx=15, pady=5, sticky="w")

        self.label_doc_nmb = ttk.Label(self.folder_frame, text='.doc文件数')
        self.label_doc_nmb.config(font=('Microsoft YaHei UI', 10))
        self.label_doc_nmb.configure(justify="center", anchor="e")
        self.label_doc_nmb.grid(row=0, column=1, padx=15, pady=5, sticky="w")

    def select_path_frame(self):
        mnplt_frame = ttk.LabelFrame(self, text="文件目录")
        mnplt_frame.grid(row=4, column=0, columnspan=2, padx=10, pady=10, ipadx=10, sticky='nsew')

        self.entry_path = ttk.Entry(mnplt_frame, justify=tk.LEFT, width=75,
                                    font=('Microsoft YaHei UI', 11))
        self.entry_path.insert(0, "浏览并选择目录")
        # self.entry_path.bind("<FocusIn>", lambda e: self.entry_path.delete('0', 'end'))
        # self.entry_path.focus_force()
        self.entry_path.grid(row=0, column=0, columnspan=2, padx=10, pady=5, sticky='nsew')

        self.browse_button = tk.Button(mnplt_frame, text="选择目录", command=self.select_directory, bg='grey',
                                       fg='white',
                                       font=('Microsoft YaHei UI', 11, 'normal'))
        self.browse_button.grid(row=1, column=0, padx=10, pady=5, ipadx=10, ipady=3)

        self.start_remove_button = tk.Button(mnplt_frame, text="开始处理", command=self.start_replace_strings_from_path,
                                             bg='purple',
                                             fg='white',
                                             font=('Microsoft YaHei UI', 11, 'bold'))
        self.start_remove_button.grid(row=1, column=1, padx=10, pady=5, ipadx=10, ipady=3)

        self.read_all_item_to_treeview_list()

        geekerWin.protocol("WM_DELETE_WINDOW", self.on_window_close)

    def author_frame(self):
        author_frame = ttk.LabelFrame(self, text="关于")
        # author_frame = ttk.Frame(self)
        author_frame.grid(row=5, column=0, columnspan=2, padx=10, pady=10, ipadx=10)

        label_link = ttk.Label(author_frame, text='www.ITGeeker.net', font=('Microsoft YaHei UI', 10), cursor="hand2")
        label_link.bind("<Button-1>", lambda e: self.open_website("https://www.itgeeker.net"))
        label_link.grid(row=0, column=0, padx=(10, 0), ipadx=10, ipady=10, sticky="w")

        label_ver = ttk.Label(author_frame, text='开源Word文字Ver 1.0.0.0', font=('Microsoft YaHei UI', 10), cursor="heart")
        label_ver.config(font=('Microsoft YaHei UI', 10))
        label_ver.bind("<Button-1>",
                       lambda e: self.open_website("https://www.itgeeker.net/itgeeker-technical-service/itgeeker_convert_doc_to_docx/"))
        label_ver.grid(row=0, column=1, padx=(10, 10), ipadx=10, ipady=10, sticky="e")


if __name__ == "__main__":
    icon_b64 = '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'
    icondata = base64.b64decode(icon_b64)
    tmp_p = tempfile.gettempdir()
    tempFile = os.path.join(tmp_p, "icon.ico")
    iconfile = open(tempFile, "wb")
    iconfile.write(icondata)
    iconfile.close()

    geekerWin = tk.Tk()
    geekerWin.wm_iconbitmap(tempFile)
    ## Delete the tempfile
    # os.remove(tempFile)

    # geekerWin.geometry("500x580")
    # geekerWin.eval('tk::PlaceWindow . center')
    window_width = 775
    window_height = 720
    display_width = geekerWin.winfo_screenwidth()
    display_height = geekerWin.winfo_screenheight()
    left = int(display_width / 2 - window_width / 2)
    top = int(display_height / 2 - window_height / 2)
    geekerWin.geometry(f'{window_width}x{window_height}+{left}+{top}')

    geekerWin.title("技术奇客小工具-文档字符替换")

    list_sheet = ListTreeSheet(geekerWin)
    list_sheet.pack()
    geekerWin.mainloop()
