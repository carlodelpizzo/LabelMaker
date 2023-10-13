import os
import shutil
import datetime
import pickle
import docx
import tkinter as tk
from tkinter import NORMAL, DISABLED, END, INSERT
from tkinter import ttk, filedialog, Toplevel, StringVar, IntVar, Checkbutton, Label, Menu

version = '1.0'

default_font = 'Arial'

letters = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm',
           'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z']
capital_letters = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M',
                   'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
numbers = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']
valid_chars = ['_', '-', ' ', *letters, *capital_letters, *numbers]


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
        self.saved_ingredients = str(self.ingredients).replace('\n', '')
        self.edited = False

    def edit_item(self, ingredients: str):
        self.edited = True
        self.ingredients = ingredients.replace('\n', '')

    def revert(self):
        self.edited = False
        self.ingredients = str(self.saved_ingredients).replace('\n', '')


class LabelMaker:
    def __init__(self):
        # Window Properties
        self.root = tk.Tk()
        self.root.geometry('800x420')
        self.root.title('Label Maker')
        self.root.resizable(False, False)

        # Program Variables
        self.username = StringVar()
        self.address = StringVar()
        self.food_items = []
        self.food_items_dict = {}
        self.selectable_items = []
        self.selected_item = None
        self.auto_save = False
        self.auto_save_name = ''
        self.label_sizes = [(4, 1)]
        self.cur_date = datetime.datetime.now().strftime('%m-%d-%Y')
        self.counter = 0

        # File management
        do_first_instance = False
        self.program_dir = f'{os.getenv("APPDATA")}/LabelMaker'
        if not os.path.isdir(self.program_dir):
            os.mkdir(self.program_dir)
            do_first_instance = True
        else:
            # Load saved information
            print('...loading saved data')
            if os.path.isfile(file_path := f'{self.program_dir}/savedata'):
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
                do_first_instance = True

        # Check for running instance
        instance_dir = f'{self.program_dir}/instance'
        if not os.path.isdir(instance_dir):
            os.mkdir(instance_dir)
        else:
            # Duplicate instance check
            print('no code for duplicate instance check')

        # Window Menu
        self.menu = Menu(self.root)
        self.file = Menu(self.menu, tearoff=0)
        self.file.add_command(label='Settings', command=self.settings_window)
        self.menu.add_cascade(label='File', menu=self.file)
        self.root.config(menu=self.menu)

        # Program UI
        self.edit_item_label = tk.Label(self.root, text='Edit Food Item', font=(default_font, 15))
        self.edit_item_label.place(x=120, y=5, anchor='n')

        self.item_name_label = tk.Label(self.root, text='Food Item:', font=(default_font, 12))
        self.item_name_label.place(x=10, y=50, anchor='w')
        self.item_name = StringVar()
        self.item_name_box = ttk.Combobox(values=self.selectable_items, postcommand=self.dropdown_opened,
                                          textvariable=self.item_name)
        self.item_name_box.bind('<<ComboboxSelected>>', self.dropdown_changed)
        self.item_name.trace('w', self.combobox_edited)
        self.item_name_box.bind('<KeyRelease>', self.combobox_user_edit)
        self.item_name_box.place(x=90, y=50, anchor='w')

        self.ingredients_label = tk.Label(self.root, text='Ingredients:', font=(default_font, 12))
        self.ingredients_label.place(x=10, y=78, anchor='w')
        self.ingredients_entry = tk.Text(self.root,  width=30, height=9, font=(default_font, 12))
        self.ingredients_entry.bind('<KeyRelease>', self.textbox_edited)
        self.ingredients_entry.place(x=10, y=90, anchor='nw')

        self.save_item_button = tk.Button(self.root, text='Save Item', width=10, command=self.save_item)
        self.save_item_button.place(x=120, y=260, anchor='ne')

        self.delete_item_button = tk.Button(self.root, text='Delete Item', width=10, command=self.delete_item)
        self.delete_item_button.place(x=122, y=260, anchor='nw')

        # Version Label
        self.version_label = Label(self.root, text=version)
        self.version_label.place(relx=1, rely=1.01, anchor='se')

        # Mainloop
        self.root.after(0, self.instance_check)
        if do_first_instance:
            self.root.after(1, self.first_instance)
        self.root.mainloop()

        # Program exit
        if any((self.username.get(), self.address.get(), self.food_items)):
            for item in self.food_items:
                item.revert()
            with open(f'{self.program_dir}/savedata', 'wb') as file:
                save_data = SaveData(self)
                pickle.dump(save_data, file)
        # Do duplicate instance check fist
        shutil.rmtree(instance_dir)

    def first_instance(self):
        # Ask to input Username and Address
        print('First Instance Ran')
        self.settings_window(first_instance=True)

    def settings_window(self, first_instance=False):
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

        username_label = tk.Label(window, text='Name:', font=(default_font, 12))
        username_label.place(x=80, y=15, anchor='e')
        username_label_entry = tk.Entry(window, textvariable=self.username, font=(default_font, 12))
        username_label_entry.place(x=82, y=17, anchor='w')

        address_label = tk.Label(window, text='Address:', font=(default_font, 12))
        address_label.place(x=80, y=45, anchor='e')
        address_label_entry = tk.Entry(window, textvariable=self.address, font=(default_font, 12))
        address_label_entry.place(x=82, y=47, anchor='w')

        if first_instance:
            window.title('Enter User Information')
            window.grab_set()
            window.focus()
            window.transient(self.root)

    def instance_check(self):
        self.root.after(150, self.instance_check)

    def save_item(self):
        if self.selected_item:
            self.selected_item.save_item()
            self.ninja_update_combobox(self.selected_item.get_name())
            self.ingredients_entry.delete('1.0', END)
            self.ingredients_entry.insert('1.0', self.selected_item.ingredients)
        else:
            new_item = FoodItem(self.item_name_box.get(), self.ingredients_entry.get('1.0', END))
            self.food_items.append(new_item)
            self.food_items_dict[new_item.name] = new_item
            self.selected_item = new_item
        self.update_combobox()
        print('Save Item Function Run')

    def delete_item(self):
        if not self.item_name_box.get():
            return
        if self.selected_item:
            self.selectable_items.pop(self.selectable_items.index(self.selected_item.get_name()))
            self.food_items.pop(self.food_items.index(self.selected_item))
            del self.food_items_dict[self.selected_item.name]
            self.selected_item = None
            self.item_name_box.set('')
            self.ingredients_entry.delete('1.0', END)
            self.update_combobox()
        print('Delete Item Function Run')

    def textbox_edited(self, event):
        if event.keysym in ['Right', 'Left', 'Up', 'Down']:
            return
        if self.selected_item:
            self.selected_item.edit_item(self.ingredients_entry.get('1.0', END))
            self.ninja_update_combobox(self.selected_item.get_name())
        self.update_combobox()
        print('Textbox edited')

    def combobox_edited(self, *_):
        if not self.item_name_box.get():
            return
        if self.selected_item:
            self.selected_item = None
            self.ingredients_entry.delete('1.0', END)
        print('Combobox edited')

    def combobox_user_edit(self, event):
        if event.keysym in ['Right', 'Left']:
            return
        value = [char for char in self.item_name_box.get() if char in valid_chars]
        value = ''.join(value)
        self.ninja_update_combobox(value)
        if item := self.food_items_dict.get(value):
            self.ingredients_entry.delete('1.0', END)
            self.ingredients_entry.insert('1.0', item.ingredients)
            self.selected_item = item
            self.ninja_update_combobox(item.get_name())
        print('User edit')

    def dropdown_changed(self, *_):
        if not (item_name := self.item_name_box.get().replace('*', '')):
            return
        if self.auto_save and self.auto_save_name != self.item_name_box.get():
            new_item = FoodItem(self.auto_save_name, self.ingredients_entry.get('1.0', END))
            self.food_items.append(new_item)
            self.food_items_dict[new_item.name] = new_item
            self.update_combobox()
            self.auto_save_name = ''
        self.auto_save = False
        if item := self.food_items_dict.get(item_name):
            self.ingredients_entry.delete('1.0', END)
            self.ingredients_entry.insert('1.0', item.ingredients)
            self.selected_item = item
        print('Dropdown Changed:', self.item_name_box.get())

    def dropdown_opened(self, *_):
        if not self.selected_item and self.item_name_box.get():
            self.auto_save = True
            self.auto_save_name = self.item_name_box.get()
        print('Dropdown Opened; Current Contents:', self.item_name_box.get())

    def update_combobox(self):
        self.selectable_items = [item.get_name() for item in self.food_items]
        self.item_name_box['values'] = self.selectable_items

    def ninja_update_combobox(self, value: str):
        selected_item = self.selected_item
        textbox_contents = self.ingredients_entry.get('1.0', END).replace('\n', '')
        self.item_name.set(value)
        self.ingredients_entry.insert('1.0', textbox_contents)
        self.selected_item = selected_item


if __name__ == '__main__':
    LabelMaker()
