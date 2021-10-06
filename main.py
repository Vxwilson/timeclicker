import ctypes
import os
import tkinter as tk
import tkinter.ttk as ttk
import tkinter.font
import tkinter.messagebox
import keyboard
import datetime
import pickle
import pyautogui
from infi.systray import SysTrayIcon
from functools import partial

import Source.scheduler as scheduler


class Application(tk.Frame):

    def __init__(self, master=None):
        super().__init__(master)
        self.data = self.load_data()
        self.settings_data = self.load_settings()
        self.scheduler = scheduler.Scheduler(root, self.start_scheduled_task, self.scheduler_cleaner)
        self.profiles = self.load_profiles()
        self.master = master
        # self.pack()

        self.terminate_loop = True
        self.terminate_click = False

        self.minimize_radio = tk.IntVar()
        self.minimize_radio.set(1 if not self.settings_data or "minimize_radio" not in self.settings_data else
                                self.settings_data["minimize_radio"])

        self.method = tk.StringVar()

        self.iteration_value = tk.IntVar()
        self.iteration_value.set(
            1 if (not self.settings_data or "iteration" not in self.settings_data) else self.settings_data["iteration"])

        # scheduler
        self.hour = tk.IntVar()
        self.hour.set(datetime.datetime.now().strftime("%H"))
        self.min = tk.IntVar()
        self.min.set(datetime.datetime.now().strftime("%M"))

        self.menubar = tk.Menu(root)
        self.filemenu = tk.Menu(self.menubar, tearoff=0)
        self.filemenu.add_command(label="Minimize", command=self.hide_window, accelerator="Ctrl+H")
        # self.filemenu.add_command(label="Save credentials", command=self.save_data, accelerator="Ctrl+S")
        self.filemenu.add_separator()
        self.filemenu.add_command(label="Preferences", command=self.open_settings_, accelerator="Alt+P")
        self.filemenu.add_separator()
        self.filemenu.add_command(label="Exit", command=root.quit)
        self.menubar.add_cascade(label="File", menu=self.filemenu)

        self.actionmenu = tk.Menu(self.menubar, tearoff=0)
        self.actionmenu.add_command(label="New task", command=lambda: self.new_profile(), accelerator="Control+N")
        self.menubar.add_cascade(label="Actions", menu=self.actionmenu)

        self.helpmenu = tk.Menu(self.menubar, tearoff=0)
        self.helpmenu.add_command(label="About...")
        self.menubar.add_cascade(label="Help", menu=self.helpmenu)
        root.config(menu=self.menubar)

        root.grid_rowconfigure(0, weight=1)
        root.grid_columnconfigure(0, weight=1)

        # # menu popup
        # self.popup = tk.Menu(self.input_text, tearoff=0)
        # self.popup.add_command(label="Add divider", command=self.add_divider, accelerator="Alt+D")
        # self.popup.add_command(label="Clear", command=self.clear_input, accelerator="Ctrl+Q")
        # self.popup.add_separator()

        # help frame
        self.help_frame = ttk.LabelFrame(text="Help", width=400)
        # self.help_frame.grid_propagate(False)
        self.help_frame.grid(row=0, column=0)
        self.help_label = ttk.Label(self.help_frame, text="""
        Cappribot version 0.1.0
        refer to GitHub readme.md for more information
        """)
        # self.help_label.grid(row=0, column=0)
        self.help_info_text = """
Timeclicker version 0.1.0 
        
To start or save a new autoclicker profile, go to Actions > New Task. Press 'Add position' button to start, then move 
your cursor to the desired position and press Control-T.
Saved profiles can be executed from Actions > Execute Profile.

To terminate an ongoing task, press Control-X. To terminate looping for an ongoing task, press Alt-X.

NOTE: for scheduled tasks to work, the window must not be closed. Instead, you can minimize it to the system tray.

Please refer to GitHub readme.md for more information.
created by vxix in 2021.
        """
        self.help_info = tk.Text(self.help_frame, width=55, height=20)
        self.help_info.insert('end', self.help_info_text)
        self.help_info.configure(state='disabled')
        self.help_info.grid(row=0, column=0, sticky="ew")

        # self.close_frame_button = ttk.Button(master=self.help_frame, text="Close",
        #                                      command=lambda: tkinterextension.lower_frame(self.help_frame))
        # # self.close_frame_button = ttk.Button(master=self.help_frame, text="Close",
        # command=lambda: tkinterextension.lower_frame(self.help_frame), style="Accentbutton")
        # self.close_frame_button.grid(row=1, column=0)

        # self.help_frame.lower()

        self.scheduled_frame = ttk.LabelFrame(text="Schedules", width=300, height=250)
        self.scheduled_frame.grid_propagate(False)
        self.scheduled_frame.grid(row=0, column=1)

        self.update_schedule()
        self.refresh_profiles()

        self.bind_keys()
        ctypes.windll.shcore.SetProcessDpiAwareness(1)  # to fix blurry text

        root.protocol('WM_DELETE_WINDOW', self.override_close)

    def bind_keys(self):
        # self.input_text.bind("<Button-3>", self.menu_popup)

        # shortcuts
        root.bind('<Control-n>', lambda e: self.new_profile())
        root.bind('<Control-h>', lambda e: self.hide_window())
        root.bind('<Control-b>', lambda e: self.stop_task())

        # self.input_text.bind('<Alt-d>', lambda e: self.add_divider())
        # self.input_text.bind('<Control-q>', lambda e: self.clear_input())

        root.bind('<Alt-p>', lambda e: self.open_settings_())

        keyboard.add_hotkey('alt + x', self.terminate_looping)
        keyboard.add_hotkey('ctrl + x', self.terminate_task)
        keyboard.add_hotkey('ctrl + shift + z', print, args='Hotkey detected')

    def open_settings_(self):
        settings = tk.Toplevel(root)
        settings.title("Settings")
        settings.minsize(550, 350)

        minimize_option_label = ttk.Label(settings, text="Close program prompt")
        minimize_option_label.grid(row=2, column=0, sticky="ewn", pady=20)
        minimize_option_radio = ttk.Radiobutton(settings, text="Always ask", variable=self.minimize_radio,
                                                value=1)
        minimize_option_radio2 = ttk.Radiobutton(settings, text="Minimize", variable=self.minimize_radio,
                                                 value=2)
        minimize_option_radio3 = ttk.Radiobutton(settings, text="Close", variable=self.minimize_radio,
                                                 value=3)
        minimize_option_radio.grid(row=2, column=0, sticky="ewn", pady=40)
        minimize_option_radio2.grid(row=2, column=0, sticky="ewn", pady=70)
        minimize_option_radio3.grid(row=2, column=0, sticky="ewn", pady=100)

        apply_button = ttk.Button(master=settings, text="Apply",
                                  command=lambda: [self.apply_settings(), settings.destroy()])
        apply_button.grid(row=4, column=0, sticky="ews", pady=10)

    def new_profile(self):
        profile_window = tk.Toplevel(root)
        profile_window.title("New profile")
        profile_window.minsize(600, 400)

        menubar = tk.Menu(profile_window)
        filemenu = tk.Menu(menubar, tearoff=0)
        # filemenu.add_command(label="Undo", command=input_text.undo, accelerator="Ctrl+Z")
        # filemenu.add_command(label="Redo", accelerator="Ctrl+Y")
        filemenu.add_command(label="Close", command=profile_window.destroy)
        menubar.add_cascade(label="File", menu=filemenu)

        actionmenu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Actions", menu=actionmenu)

        profile_window.config(menu=menubar)

        # login_frame = ttk.LabelFrame(profile_window, text="Recipient Details")
        # login_frame.grid(row=0, column=1, sticky="new")
        # login_frame.grid_rowconfigure([0], weight=1, minsize=50)
        # login_frame.grid_columnconfigure([0], weight=1, minsize=130)

        # menu popup
        popup = tk.Menu(tearoff=0)

        def menu_popup(event):
            try:
                popup.tk_popup(event.x_root, event.y_root, 0)
            finally:
                popup.grab_release()

        popup.add_separator()

        selection_frame = ttk.LabelFrame(profile_window, text="Add positions")
        selection_frame.grid(row=0, column=0, sticky="new")
        selection_frame.grid_rowconfigure([0, 1], weight=1)
        selection_frame.grid_columnconfigure([0, 1], weight=1)

        pos_idx = 0
        positions = []
        delays = []

        def get_point():
            nonlocal pos_idx, positions, delays

            pos = pyautogui.position()
            ttk.Label(selection_frame, text=f"x: {pos.x} y:{pos.y}  Delay:") \
                .grid(row=pos_idx + 1, column=0, sticky="new")

            delay = tk.IntVar()
            delay.set(1001)

            delays.append(delay)
            positions.append([pos.x, pos.y])

            ttk.Spinbox(selection_frame, width=8, from_=1, to=100000, increment=500, wrap=True, textvariable=delays[pos_idx]) \
                .grid(row=pos_idx + 1, column=1, sticky="new")

            pos_idx += 1

        add_pos_button = ttk.Button(selection_frame, text="Add", command=lambda: [])
        add_pos_button.grid(row=0, column=0)

        add_pos_button.bind("<Control-t>", lambda e: get_point())
        # add_pos_button.bind("<Button-1>", lambda e: get_point())

        technical_frame = ttk.LabelFrame(profile_window, text="Options")
        technical_frame.grid(row=0, column=1, sticky="new")
        technical_frame.grid_rowconfigure([0, 1, 2], weight=1)
        technical_frame.grid_columnconfigure([0, 1], weight=1)

        task_label_var = tk.StringVar()
        task_label_var.set("label")  # todo add label index automatically
        task_label = ttk.Entry(technical_frame, textvariable=task_label_var)
        task_label.grid(row=0, column=0, sticky="new")

        start_idx = tk.IntVar()
        start_idx.set(1)

        start_idx_label = ttk.Label(technical_frame, text="Starting index:")
        start_idx_label.grid(row=1, column=0, sticky="ews")
        start_idx_selector = ttk.Spinbox(technical_frame, width=5, from_=1, wrap=True, textvariable=start_idx)
        start_idx_selector.grid(row=1, column=0, sticky="ews", padx=120)

        loop = tk.BooleanVar()
        loop.set(True)

        repeat_checkbox = ttk.Checkbutton(technical_frame, text="Loop", variable=loop)
        repeat_checkbox.grid(row=2, column=1, sticky="w")
        # todo show iteration spinbox after 'loop' is disabled
        #
        # iteration_label = ttk.Label(master=technical_frame, text="Iteration:")
        # iteration_label.grid(row=1, column=0, sticky="w", pady=15)
        # iteration_box = ttk.Spinbox(master=technical_frame, width=4, from_=1, to=100, wrap=True,
        #                             textvariable=self.iteration_value)
        # iteration_box.grid(row=1, column=1, sticky="w")

        schedule_frame = ttk.LabelFrame(profile_window, text="Schedule")
        schedule_frame.grid(row=1, column=1, sticky="new")

        scheduler_label = ttk.Label(master=schedule_frame, text="Schedule:")
        scheduler_label.grid(row=2, column=0, sticky="w")
        # todo check possible bug that occur when 00 is passed instead of 0
        scheduler_hour = ttk.Spinbox(master=schedule_frame, width=5, from_=0, to=23, increment=1,
                                     textvariable=self.hour, wrap=True)
        scheduler_hour.grid(row=2, column=1, sticky="w", padx=0)
        scheduler_minute = ttk.Spinbox(master=schedule_frame, width=5, from_=0, to=59, increment=1,
                                       textvariable=self.min, wrap=True)
        scheduler_minute.grid(row=2, column=1, sticky="w", padx=60)
        scheduler_hourlabel = ttk.Label(schedule_frame, text="Hours")
        scheduler_hourlabel.grid(row=2, column=1, sticky="w", padx=120)

        profile_window.grid_rowconfigure(0, weight=1)
        profile_window.grid_columnconfigure(0, weight=1)

        def add_profile(pos_details, month=1, day=1, hour=12, minute=0, second=0):

            delay_get = [x.get() for x in delays]

            if self.profiles:
                profiles = self.profiles
                profiles.append(
                    {'label': task_label_var.get(), 'pos_details': pos_details, 'delays': delay_get, 'month': month,
                     'day': day, 'hour': hour, 'minute': minute, 'second': second})
                data = {'entry': profiles}
            else:
                data = {'entry': [
                    {'label': task_label_var.get(), 'pos_details': pos_details, 'delays': delay_get, 'month': month,
                     'day': day, 'hour': hour, 'minute': minute,
                     'second': second}]}
            with open('Source/Resources/profiles.txt', 'wb') as file:
                pickle.dump(data, file)
            self.profiles = self.load_profiles()
            self.refresh_profiles()

        # todo add instant start button
        start_button = ttk.Button(master=profile_window, text="Start schedule",
                                  command=lambda: [add_profile(pos_details=positions, hour=scheduler_hour.get(),
                                                               minute=scheduler_minute.get()),
                                                   self.scheduler.add_schedule(idx=len(self.profiles) - 1,
                                                                               loop=loop.get(),
                                                                               start_idx=start_idx.get(),
                                                                               minute=self.min.get(),
                                                                               hour=self.hour.get()),
                                                   self.update_schedule(), profile_window.destroy()])
        start_button.grid(row=2, column=0, sticky="ews")
        apply_button = ttk.Button(master=profile_window, text="Save profile",
                                  command=lambda: [
                                      add_profile(pos_details=positions, hour=scheduler_hour.get(),
                                                  minute=scheduler_minute.get())])
        apply_button.grid(row=2, column=1, sticky="ews")

        # input_text.bind("<Button-3>", menu_popup)

    def update_schedule(self):

        def remover(idx):
            self.scheduler.remove_schedule(idx)
            self.update_schedule()

        for item in self.scheduled_frame.winfo_children():
            item.destroy()

        schedules = self.scheduler.load_schedules()

        for idx, schedule in enumerate(schedules):
            # ttk.Label(self.scheduled_frame, text=f'{schedule["ref_label"]}').grid(row=idx, column=0)
            ttk.Label(self.scheduled_frame, text=f'{schedule["hour"]}:{schedule["minute"]} hours').grid(row=idx,
                                                                                                        column=0)
            ttk.Button(self.scheduled_frame, text="Remove",
                       command=partial(remover, idx)) \
                .grid(row=idx, column=1, padx=25, pady=5, sticky="e")

    def refresh_profiles(self):
        try:
            self.actionmenu.delete("Execute profile")
        except:
            pass

        if self.profiles and bool(self.profiles):
            nested_menu = tk.Menu(self.actionmenu)

            for index, profile in enumerate(self.profiles):
                nested_menu.add_command(label=profile["label"], command=partial(self.open_profile, index))
            # nested_menu.add_command(label="See all", command=list_links) todo add see all link function
            self.actionmenu.add_cascade(label="Execute profile", menu=nested_menu)
        else:
            pass

    def open_profile(self, index):
        profile_window = tk.Toplevel(root)
        profile_window.title("Start auto clicker")
        profile_window.minsize(400, 250)

        menubar = tk.Menu(profile_window)
        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label="Close", command=profile_window.destroy)
        menubar.add_cascade(label="File", menu=filemenu)

        actionmenu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Actions", menu=actionmenu)

        profile_window.config(menu=menubar)

        # selection_frame = ttk.LabelFrame(profile_window, text="Add positions")
        # selection_frame.grid(row=0, column=0, sticky="new")
        # selection_frame.grid_rowconfigure([0, 1], weight=1)
        # selection_frame.grid_columnconfigure([0, 1], weight=1)

        profile = self.load_profiles(index)

        start_idx = tk.IntVar()
        start_idx.set(1)

        loop = tk.BooleanVar()
        loop.set(False)

        start_idx_label = ttk.Label(profile_window, text="Starting index:")
        start_idx_label.grid(row=0, column=0, sticky="ews")
        start_idx_selector = ttk.Spinbox(profile_window, width=5, from_=1,
                                         to=len(profile["pos_details"]), wrap=True, textvariable=start_idx)
        start_idx_selector.grid(row=0, column=0, sticky="ews", padx=120)

        loop_checkbox = ttk.Checkbutton(profile_window, text="Loop", variable=loop)
        loop_checkbox.grid(row=1, column=0, sticky="ews")

        schedule_frame = ttk.LabelFrame(profile_window, text="Schedule")
        schedule_frame.grid(row=1, column=1, sticky="new")

        scheduler_label = ttk.Label(master=schedule_frame, text="Schedule:")
        scheduler_label.grid(row=2, column=0, sticky="w")
        # todo check possible bug that occur when 00 is passed instead of 0
        scheduler_hour = ttk.Spinbox(master=schedule_frame, width=5, from_=0, to=23, increment=1,
                                     textvariable=self.hour, wrap=True)
        scheduler_hour.grid(row=2, column=1, sticky="w", padx=0)
        scheduler_minute = ttk.Spinbox(master=schedule_frame, width=5, from_=0, to=59, increment=1,
                                       textvariable=self.min, wrap=True)
        scheduler_minute.grid(row=2, column=1, sticky="w", padx=60)
        scheduler_hourlabel = ttk.Label(schedule_frame, text="Hours")
        scheduler_hourlabel.grid(row=2, column=1, sticky="w", padx=120)

        start_button = ttk.Button(master=profile_window, text="Execute instantly",
                                  command=lambda: [self.start_profile(index, start_idx.get() - 1, loop.get())])
        start_button.grid(row=2, column=0, sticky="ews")
        schedule_button = ttk.Button(master=profile_window, text="Schedule task",
                                     command=lambda: [self.scheduler.add_schedule(idx=index,
                                                                                  loop=loop.get(),
                                                                                  start_idx=start_idx.get() - 1,
                                                                                  hour=self.hour.get(),
                                                                                  minute=self.min.get()),
                                                      self.update_schedule(), profile_window.destroy()])
        schedule_button.grid(row=2, column=0, sticky="ews")

        delete_button = ttk.Button(master=profile_window, text="Delete profile",
                                   command=lambda: [self.remove_profile(index), profile_window.destroy()])
        delete_button.grid(row=2, column=1, sticky="ews")

        profile_window.grid_rowconfigure([0, 1, 2], weight=1)
        profile_window.grid_columnconfigure([0, 1], weight=1)

    def start_profile(self, index, start_index=0, loop=False):
        profile = self.load_profiles(index)
        delays = profile["delays"]
        positions = profile["pos_details"]

        self.terminate_loop = not loop
        self.terminate_click = False
        n = 1  # used to make sure start index check is only triggered first loop
        while True:
            # todo fix not responding bug
            for idx, pos in enumerate(positions):
                if self.terminate_click is True:
                    return True
                if n > 1 or idx >= start_index:
                    pyautogui.moveTo(positions[idx][0], positions[idx][1])
                    root.after(5)
                    pyautogui.click()
                    root.after(delays[idx])
            if self.terminate_loop is True:
                return True
            n += 1

    def terminate_task(self):
        print('Task terminated')
        self.terminate_click = True

    def terminate_looping(self):
        print('Looping terminated')
        self.terminate_loop = True

    def stop_task(self):
        pass

    def start_scheduled_task(self, index, **kwargs):  # click stuff
        schedule = self.scheduler.load_schedules(index)
        self.start_profile(index=schedule["index"], loop=schedule["loop"], start_index=schedule["start_idx"])

    # used by scheduler to update schedule after removing schedule
    def scheduler_cleaner(self):
        self.update_schedule()

    def load_profiles(self, idx=-1):
        if os.path.exists('Source/Resources/profiles.txt'):
            try:
                with open('Source/Resources/profiles.txt', 'r+b') as file:
                    schedules = pickle.load(file)["entry"]
                    if idx != -1 and len(schedules) > idx >= 0:
                        return schedules[idx]
                    else:
                        return schedules
            except (EOFError, KeyError) as e:
                print(e)
                return {}
        else:
            return {}

    def remove_profile(self, idx):
        if os.path.exists('Source/Resources/profiles.txt'):
            try:
                with open('Source/Resources/profiles.txt', 'r+b') as file:
                    profiles = pickle.load(file)["entry"]
                    if len(profiles) > idx >= 0:
                        del profiles[idx]

                    if len(profiles) == 0:  # profile now empty
                        data = {}
                    else:
                        data = {'entry': profiles}
                    with open('Source/Resources/profiles.txt', 'wb') as file:
                        pickle.dump(data, file)

                    self.profiles = self.load_profiles()
                    self.refresh_profiles()
            except (EOFError, KeyError) as e:
                print(e)
                return {}
        else:
            return {}

    # shows exit prompt when users press 'X' button
    def override_close(self):
        if self.minimize_radio.get() == 2:
            self.hide_window()
        elif self.minimize_radio.get() == 3:
            root.destroy()
        else:
            res = tk.messagebox.askyesnocancel('Close Window', 'Minimize window '
                                                               'instead of closing? This is essential for '
                                                               'scheduled tasks to happen. (Changable in Setings)')
            if res:
                self.hide_window()
                pass
            elif res is False:
                root.destroy()
                pass
            elif res is None:
                pass

    def hide_window(self):
        def clicked(icon=None):
            self.show_window()
            try:
                systray.shutdown()
            except:
                print()
                pass

        root.withdraw()
        menu_options = (("Show window", None, clicked),)
        systray = SysTrayIcon("Source/Resources/Icon/picturexviewer.ico", "Cappribot", menu_options,
                              default_menu_index=0)
        systray.start()

    def show_window(self):
        root.deiconify()

    def load_settings(self):
        if os.path.exists('Source/Resources/settings.txt'):
            try:
                with open('Source/Resources/settings.txt', 'r+b') as file:
                    return pickle.load(file)
            except EOFError:
                return {}
        else:
            return {}

    def apply_settings(self):
        self.save_settings()

    def save_settings(self):
        data = {
            'minimize_radio': self.minimize_radio.get()
        }
        # 'method': self.send_method.current()}
        with open('Source/Resources/settings.txt', 'wb') as file:
            pickle.dump(data, file)

    def load_data(self):
        if os.path.exists('Source/Resources/save.txt'):
            try:
                with open('Source/Resources/save.txt', 'r+b') as file:
                    return pickle.load(file)
            except EOFError:
                return {}


root = tk.Tk()
# ttk.Style().configure("TButton", padding=6, relief="flat", foreground="#E8E8E8", background="#292929")
default_font = tk.font.nametofont("TkDefaultFont")
# print(tk.font.families())
# default_font.configure(family="Garamond", size=13)
root.tk.call('source', 'Source/Resources/Style/azure.tcl')
root.tk.call("set_theme", "dark")
default_font.configure(size=11)
# root.geometry("1050x600")
root.title("Timeclicker v0.1.0a")
root.iconphoto(False, tk.PhotoImage(file='Source/Resources/Icon/gradient_less_saturated.png'))
root.resizable(False, False)
app = Application(master=root)
app.mainloop()
