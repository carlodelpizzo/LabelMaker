import os
import shutil
import datetime
import pickle
import tkinter as tk
import psutil
from tkinter import ttk, filedialog, Toplevel, StringVar, Label, Menu, END, DISABLED, NORMAL
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_TAB_ALIGNMENT, WD_BREAK

version = '1.0'

program_font = 'Arial'

program_dir = f'{os.getenv("APPDATA")}/LabelMaker'
instance_dir = f'{program_dir}/instance'

letters = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm',
           'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z']
capital_letters = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M',
                   'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
numbers = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']
valid_chars = ['(', ')', '_', '-', ' ', *letters, *capital_letters, *numbers]


class SaveData:
    def __init__(self, label_maker: object):
        self.username_str = label_maker.username.get()
        self.address_str = label_maker.address.get()
        self.food_items = label_maker.food_items


class FoodItem:
    def __init__(self, name: str, ingredients: str):
        self.name = name
        self.ingredients = ingredients
        self.saved_ingredients = ingredients
        self.edited = False

    def get_name(self):
        return f'*{self.name}' if self.edited else self.name

    def save_item(self):
        self.ingredients = self.ingredients.replace('\n', '')
        self.saved_ingredients = str(self.ingredients)
        self.edited = False

    def edit_item(self, ingredients: str):
        self.ingredients = ingredients.replace('\n', '')
        self.edited = (self.ingredients != self.saved_ingredients)

    def revert(self):
        self.edited = False
        self.ingredients = str(self.saved_ingredients).replace('\n', '')


class LabelMaker:
    def __init__(self):
        # Check for running instance
        running_processes = [process.name() for process in psutil.process_iter()]
        if running_processes.count('LabelMaker.exe') > 2:
            return

        # Window Properties
        self.root = tk.Tk()
        self.root.geometry('670x310')
        self.root.title('Label Maker')
        self.root.resizable(False, False)

        # Program Variables
        self.username = StringVar()
        self.address = StringVar()
        self.food_items = []
        self.food_items_dict = {}
        self.selectable_items = []
        self.selected_item = None
        self.label_size = (4, 1)
        self.labels_to_print = {}
        self.auto_save = False
        self.auto_save_name = ''
        self.date_stamp = datetime.datetime.now().strftime('%m-%d-%Y')
        self.save_path = f'{os.path.join(os.environ["USERPROFILE"])}/Desktop'
        self.counter = 0

        # File management
        first_run = False
        if not os.path.isdir(program_dir):
            os.mkdir(program_dir)
            first_run = True
        elif os.path.isfile(file_path := f'{program_dir}/savedata'):
            # Load saved information
            with open(file_path, 'rb') as file:
                save_data = pickle.load(file)
            if type(save_data) is not SaveData:
                raise TypeError
            self.username.set(save_data.username_str)
            self.address.set(save_data.address_str)
            self.food_items = save_data.food_items
            self.food_items_dict = {item.name: item for item in self.food_items}
            self.selectable_items = [item.get_name() for item in self.food_items]
        else:
            first_run = True

        # Window Menu
        self.menu = Menu(self.root)
        self.file = Menu(self.menu, tearoff=0)
        self.file.add_command(label='Settings', command=self.settings_window)
        self.menu.add_cascade(label='File', menu=self.file)
        self.root.config(menu=self.menu)

        # Program UI
        self.edit_item_label = tk.Label(self.root, text='Edit Food Item', font=(program_font, 15))
        self.edit_item_label.place(x=147, y=2, anchor='n')

        self.item_name_label = tk.Label(self.root, text='Food Item:', font=(program_font, 12))
        self.item_name_label.place(x=10, y=50, anchor='w')
        self.item_name = StringVar()
        self.item_name_box = ttk.Combobox(values=self.selectable_items, postcommand=self.dropdown_opened,
                                          textvariable=self.item_name)
        self.item_name_box.bind('<<ComboboxSelected>>', self.dropdown_changed)
        self.item_name.trace('w', self.combobox_edited)
        self.item_name_box.bind('<KeyRelease>', self.combobox_user_edit)
        self.item_name_box.place(x=90, y=50, anchor='w')

        self.edit_name_button = tk.Button(self.root, text='\u270E',
                                          command=lambda: self.edit_item_name(food_item=self.selected_item))
        self.edit_name_button.place(x=240, y=50, anchor='w')

        self.ingredients_label = tk.Label(self.root, text='Ingredients:', font=(program_font, 12))
        self.ingredients_label.place(x=10, y=78, anchor='w')
        self.ingredients_entry = tk.Text(self.root, width=30, height=9, font=(program_font, 12))
        self.ingredients_entry.bind('<KeyRelease>', self.textbox_edited)
        self.ingredients_entry.place(x=10, y=90, anchor='nw')

        self.save_item_button = tk.Button(self.root, text='Save Item', width=10, command=self.save_item)
        self.save_item_button.place(x=146, y=260, anchor='ne')

        self.delete_item_button = tk.Button(self.root, text='Delete Item', width=10, command=self.delete_item)
        self.delete_item_button.place(x=148, y=260, anchor='nw')

        self.separator = ttk.Separator(self.root, orient='vertical')
        self.separator.place(x=295, y=0, relwidth=0.2, relheight=1)

        self.selected_label = tk.Label(self.root, text='Labels to print:', font=(program_font, 15))
        self.selected_label.place(x=480, y=2, anchor='n')

        self.selected_item_name = StringVar()
        self.selected_label = tk.Label(self.root, textvariable=self.selected_item_name, font=(program_font, 18))
        self.selected_label.place(x=495, y=52, anchor='e')

        self.spinbox = tk.Spinbox(self.root, from_=0, to=99, width=3)
        self.spinbox.place(x=505, y=52, anchor='w')
        self.spinbox.delete(0, END)
        self.spinbox.insert(0, '1')
        self.spinbox['state'] = 'readonly'

        self.add_item_button = tk.Button(self.root, text='Add', width=3, command=self.add_item)
        self.add_item_button.place(x=540, y=52, anchor='w')

        self.items_to_print_label = tk.Label(self.root, text='Labels to print:', font=(program_font, 12))
        self.items_to_print_label.place(x=393, y=95, anchor='s')
        self.items_to_print_entry = tk.Text(self.root, width=30, height=9, font=(program_font, 12))
        self.items_to_print_entry.place(x=480, y=96, anchor='n')
        self.items_to_print_entry['state'] = DISABLED

        self.create_labels_button = tk.Button(self.root, text='Create Labels', width=10, command=self.save_labels)
        self.create_labels_button.place(x=480, y=270, anchor='n')

        # Version Label
        self.version_label = Label(self.root, text=version)
        self.version_label.place(relx=1, rely=1.01, anchor='se')

        # Mainloop
        if first_run:
            self.root.after(1, self.on_first_run)
        self.root.protocol('WM_DELETE_WINDOW', self.on_program_exit)
        self.root.mainloop()

    def on_program_exit(self):
        if any(item.edited for item in self.food_items):
            self.save_changes()
            return
        with open(f'{program_dir}/savedata', 'wb') as file:
            save_data = SaveData(self)
            pickle.dump(save_data, file)
        self.root.destroy()

    def on_first_run(self):
        # Ask to input Username and Address
        self.settings_window(first_run=True)

    def settings_window(self, first_run=False):
        def blink_entry(entry_to_blink=None):
            if not self.counter:
                return
            if not entry_to_blink:
                self.counter = 0
                return
            if 'username' in entry_to_blink:
                if username_label_entry['background'] != 'SystemWindow':
                    username_label_entry['background'] = 'SystemWindow'
                else:
                    username_label_entry['background'] = 'pink'
                username_label_entry.update()
            if 'address' in entry_to_blink:
                if address_label_entry['background'] != 'SystemWindow':
                    address_label_entry.config(background='SystemWindow')
                    address_label_entry['background'] = 'SystemWindow'
                else:
                    address_label_entry['background'] = 'pink'
                address_label_entry.update()
            self.counter -= 1
            window.after(150, blink_entry, entry_to_blink)

        def window_close():
            empty_entries = []
            if not self.username.get():
                empty_entries.append('username')
            if not self.address.get():
                empty_entries.append('address')
            if empty_entries:
                self.counter = 6
                blink_entry(empty_entries)
            else:
                window.destroy()
        window = Toplevel(self.root)
        window.focus()
        window.title('Edit Settings')
        window.geometry('275x250')
        window.geometry(f'+{self.root.winfo_rootx()}+{self.root.winfo_rooty()}')
        window.resizable(False, False)
        window.protocol('WM_DELETE_WINDOW', window_close)

        username_label = tk.Label(window, text='Name:', font=(program_font, 12))
        username_label.place(x=80, y=15, anchor='e')
        username_label_entry = tk.Entry(window, textvariable=self.username, font=(program_font, 12))
        username_label_entry.place(x=82, y=17, anchor='w')

        address_label = tk.Label(window, text='Address:', font=(program_font, 12))
        address_label.place(x=80, y=45, anchor='e')
        address_label_entry = tk.Entry(window, textvariable=self.address, font=(program_font, 12))
        address_label_entry.place(x=82, y=47, anchor='w')

        if first_run:
            window.title('Enter User Information')
            window.grab_set()
            window.focus()
            window.transient(self.root)

    def save_item(self):
        if not self.item_name_box.get():
            return
        if self.selected_item:
            self.selected_item.save_item()
            self.update_combobox_text_only(self.selected_item.get_name())
            self.ingredients_entry.delete('1.0', END)
            self.ingredients_entry.insert('1.0', self.selected_item.ingredients.replace('\n', ''))
        else:
            new_item = FoodItem(self.item_name_box.get(), self.ingredients_entry.get('1.0', END))
            self.food_items.append(new_item)
            self.food_items_dict[new_item.name] = new_item
            self.change_selected_item(new_item)
        self.update_combobox()

    def delete_item(self):
        def window_close(delete=False):
            if delete and self.selected_item:
                if self.selected_item in self.labels_to_print:
                    del self.labels_to_print[self.selected_item]
                    self.update_items_to_print_entry()
                self.selectable_items.pop(self.selectable_items.index(self.selected_item.get_name()))
                self.food_items.pop(self.food_items.index(self.selected_item))
                del self.food_items_dict[self.selected_item.name]
                self.change_selected_item()
                self.item_name_box.set('')
                self.ingredients_entry.delete('1.0', END)
                self.update_combobox()
            del_window.destroy()

        if not self.item_name_box.get():
            return

        del_window = Toplevel(self.root)
        del_window.grab_set()
        del_window.focus()
        del_window.transient(self.root)
        del_window.title('Confirm Delete')
        del_window.geometry('240x80')
        del_window.geometry(f'+{self.root.winfo_rootx()}+{self.root.winfo_rooty()}')
        del_window.resizable(False, False)
        del_window.protocol('WM_DELETE_WINDOW', window_close)

        item_name_label = tk.Label(del_window, text='Delete Item?', font=(program_font, 12))
        item_name_label.place(x=120, y=22, anchor='s')

        save_button = tk.Button(del_window, text='Delete', width=10, command=lambda: window_close(delete=True))
        save_button.place(x=115, y=32, anchor='ne')

        cancel_button = tk.Button(del_window, text='Don\'t Delete', width=10, command=window_close)
        cancel_button.place(x=125, y=32, anchor='nw')

    def edit_item_name(self, food_item):
        def window_close():
            def cancel_edit(*event):
                if event and event[0].keysym != 'Return':
                    return
                error_window.destroy()
                window.destroy()

            if food_item.name != new_item_name.get():
                if new_item_name.get() in self.food_items_dict:
                    error_window = Toplevel(window)
                    error_window.grab_set()
                    error_window.focus()
                    error_window.transient(window)
                    error_window.title('Error')
                    error_window.geometry('225x100')
                    error_window.geometry(f'+{window.winfo_rootx()}+{window.winfo_rooty()}')
                    error_window.resizable(False, False)

                    existing_name_label = tk.Label(error_window, text='Name already exists', font=(program_font, 12))
                    existing_name_label.place(x=112, y=29, anchor='s')

                    edit_name_button = tk.Button(error_window, text='Cancel Edit', command=cancel_edit)
                    edit_name_button.bind('<KeyRelease>', cancel_edit)
                    edit_name_button.focus_set()
                    edit_name_button.place(x=112, y=31, anchor='n')
                    return
                del self.food_items_dict[food_item.name]
                food_item.name = new_item_name.get()
                self.food_items_dict[food_item.name] = food_item
                self.update_combobox_text_only(food_item.name)
                self.update_combobox()
                self.update_items_to_print_entry()
            window.destroy()

        def key_release(event):
            if event.keysym == 'Return':
                window_close()

        if not food_item:
            return

        new_item_name = StringVar()
        new_item_name.set(food_item.name)

        window = Toplevel(self.root)
        window.grab_set()
        window.focus()
        window.transient(self.root)
        window.title('Edit Item Name')
        window.geometry('225x100')
        window.geometry(f'+{self.root.winfo_rootx()}+{self.root.winfo_rooty()}')
        window.resizable(False, False)
        window.protocol('WM_DELETE_WINDOW', window_close)

        item_name_label = tk.Label(window, text='Item Name:', font=(program_font, 12))
        item_name_label.place(x=112, y=29, anchor='s')
        item_name_entry = tk.Entry(window, textvariable=new_item_name, font=(program_font, 12))
        item_name_entry.bind('<KeyRelease>', key_release)
        item_name_entry.place(x=112, y=31, anchor='n')

    def save_changes(self):
        def window_close(do='cancel'):
            if do == 'cancel':
                window.destroy()
                return
            if do == 'save':
                for item in self.food_items:
                    item.save_item()
            else:
                for item in self.food_items:
                    item.revert()
            self.on_program_exit()

        window = Toplevel(self.root)
        window.grab_set()
        window.focus()
        window.transient(self.root)
        window.title('Unsaved Changes')
        window.geometry('240x80')
        window.geometry(f'+{self.root.winfo_rootx()}+{self.root.winfo_rooty()}')
        window.resizable(False, False)
        window.protocol('WM_DELETE_WINDOW', window_close)

        item_name_label = tk.Label(window, text='Save unsaved changes?', font=(program_font, 12))
        item_name_label.place(x=120, y=22, anchor='s')

        save_button = tk.Button(window, text='Save', width=10, command=lambda: window_close(do='save'))
        save_button.place(x=115, y=32, anchor='ne')

        cancel_button = tk.Button(window, text='Don\'t Save', width=10, command=lambda: window_close(do='dont save'))
        cancel_button.place(x=125, y=32, anchor='nw')

    def textbox_edited(self, event):
        if event.keysym in ['Right', 'Left', 'Up', 'Down']:
            return
        if event.keysym == 'Return':
            self.save_item()
            textbox_contents = self.ingredients_entry.get('1.0', END).replace('\n', '')
            self.ingredients_entry.delete('1.0', END)
            self.ingredients_entry.insert('1.0', textbox_contents)
            return
        if self.selected_item:
            self.selected_item.edit_item(self.ingredients_entry.get('1.0', END))
            self.update_combobox_text_only(self.selected_item.get_name())
        self.update_combobox()

    def combobox_edited(self, *_):
        if not self.item_name_box.get():
            return
        if self.selected_item:
            self.change_selected_item()
            self.ingredients_entry.delete('1.0', END)

    def combobox_user_edit(self, event):
        if event.keysym in ['Right', 'Left']:
            return
        value = [char for char in self.item_name_box.get() if char in valid_chars]
        value = ''.join(value).lstrip()
        self.update_combobox_text_only(value)
        if item := self.food_items_dict.get(value):
            self.ingredients_entry.delete('1.0', END)
            self.ingredients_entry.insert('1.0', item.ingredients.replace('\n', ''))
            self.change_selected_item(item)
            self.update_combobox_text_only(item.get_name())

    def dropdown_changed(self, *_):
        if not (item_name := self.item_name_box.get().replace('*', '')):
            return
        if self.auto_save and self.auto_save_name != self.item_name_box.get():
            new_item = FoodItem(self.auto_save_name, self.ingredients_entry.get('1.0', END))
            new_item.edited = True
            self.food_items.append(new_item)
            self.food_items_dict[new_item.name] = new_item
            self.update_combobox()
            self.auto_save_name = ''
        self.auto_save = False
        if item := self.food_items_dict.get(item_name):
            self.ingredients_entry.delete('1.0', END)
            self.ingredients_entry.insert('1.0', item.ingredients.replace('\n', ''))
            self.change_selected_item(item)

    def dropdown_opened(self, *_):
        if not self.selected_item and self.item_name_box.get():
            self.auto_save = True
            self.auto_save_name = self.item_name_box.get()

    def update_combobox(self):
        self.selectable_items = [item.get_name() for item in self.food_items]
        self.item_name_box['values'] = self.selectable_items

    def update_combobox_text_only(self, value: str):
        selected_item = self.selected_item
        textbox_contents = self.ingredients_entry.get('1.0', END).replace('\n', '')
        self.item_name.set(value)
        self.ingredients_entry.delete('1.0', END)
        self.ingredients_entry.insert('1.0', textbox_contents)
        self.change_selected_item(selected_item)

    def add_item(self):
        if not self.selected_item:
            return
        if self.spinbox.get() == '0' and self.labels_to_print.get(self.selected_item):
            del self.labels_to_print[self.selected_item]
        elif self.spinbox.get() == '0':
            return
        else:
            self.labels_to_print[self.selected_item] = self.spinbox.get()
        self.update_items_to_print_entry()

    def update_items_to_print_entry(self):
        labels_to_print = []
        for item in self.labels_to_print:
            labels_to_print.extend([f'{item.name} ({self.labels_to_print[item]})', '\n'])
        labels_to_print = ''.join(labels_to_print[:-1])
        self.items_to_print_entry['state'] = NORMAL
        self.items_to_print_entry.delete('1.0', END)
        self.items_to_print_entry.insert('1.0', labels_to_print)
        self.items_to_print_entry['state'] = DISABLED

    def create_labels(self):
        if not self.labels_to_print:
            return None

        # Format Document
        document = Document()
        document.styles['Normal'].paragraph_format.space_before = Pt(0)
        document.styles['Normal'].paragraph_format.space_after = Pt(0)
        document.styles['Normal'].font.name = 'Arial'
        for section in document.sections:
            section.page_width = Inches(self.label_size[0])
            section.page_height = Inches(self.label_size[1])
            margin_size = Inches(1/16)
            section.top_margin = margin_size
            section.bottom_margin = margin_size
            section.left_margin = margin_size
            section.right_margin = margin_size

        # Create document contents
        for i, item in enumerate(self.labels_to_print):
            for j in range(limit := int(self.labels_to_print[item])):
                para1 = document.add_paragraph()
                item_name = para1.add_run(f'{item.name}')
                item_name.bold = True
                item_name.font.size = Pt(11)
                para1.add_run(f'\t{self.username.get()} - {self.address.get()}').font.size = Pt(10)
                margin_end = Inches((sec := document.sections[0]).page_width.inches -
                                    (sec.left_margin.inches + sec.right_margin.inches))
                tab_stops = para1.paragraph_format.tab_stops
                tab_stops.add_tab_stop(margin_end, WD_TAB_ALIGNMENT.RIGHT)

                para2 = document.add_paragraph()
                ingredients = para2.add_run('Ingredients: ')
                ingredients.bold = True
                ingredients.font.size = Pt(9)
                para2.add_run(item.ingredients.replace('\n', '')).font.size = Pt(8)
                para3 = document.add_paragraph()
                date = para3.add_run('Date: ')
                date.font.size = Pt(8)
                date_stamp = para3.add_run('')  # Automated date stamp goes here
                date_stamp.font.size = Pt(8)
                date_stamp.bold = True
                weight_spacing = ''.join([' ' for _ in range(40)])
                weight = para3.add_run(f'{weight_spacing}Weight:')
                weight.font.size = Pt(8)
                if i == len(self.labels_to_print) - 1 and j == limit - 1:
                    continue
                weight.add_break(WD_BREAK.PAGE)

        return document

    def save_labels(self):
        if not self.labels_to_print:
            return
        out_file = filedialog.asksaveasfile(initialdir=self.save_path, initialfile=f'Labels {self.date_stamp}.docx',
                                            filetypes=[('Word Document', '*.docx')])
        if not out_file:
            return
        if not (out_file_path := out_file.name).endswith('.docx'):
            out_file.close()
            os.rename(out_file_path, (out_file_path := f'{out_file_path}.docx'))
        self.save_path = os.path.dirname(out_file_path)
        document = self.create_labels()
        document.save(out_file_path)

    def change_selected_item(self, item=None):
        item_name = item.name if item else ''
        self.selected_item = item
        self.selected_item_name.set(item_name)
        spinbox_text = '1'
        if self.selected_item in self.labels_to_print:
            spinbox_text = self.labels_to_print[self.selected_item]
        self.spinbox['state'] = NORMAL
        self.spinbox.delete(0, END)
        self.spinbox.insert(0, spinbox_text)
        self.spinbox['state'] = 'readonly'


if __name__ == '__main__':
    LabelMaker()
