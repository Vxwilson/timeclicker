import os
import pickle
import datetime


class Scheduler:

    def __init__(self, root, func, cleaner=None):
        self.process = []
        self.schedules = self.load_schedules()
        self.root = root
        self.func = func
        self.cleaner = cleaner
        self.recheck_all_schedules()

    def add_schedule(self, idx, start_idx=0, loop=False, month=1, day=1, hour=12, minute=0, second=0):
        if self.schedules:

            schedules = self.schedules
            schedules.append({'index': idx, 'start_idx': start_idx, 'loop': loop, 'month': month, 'day': day,
                              'hour': hour, 'minute': minute, 'second': second})
            data = {'entry': schedules}
        else:
            data = {'entry': [{'index': idx, 'start_idx': start_idx, 'loop': loop, 'month': month, 'day': day,
                               'hour': hour, 'minute': minute, 'second': second}]}

        with open('Source/Resources/schedules.txt', 'wb') as file:
            pickle.dump(data, file)
        self.schedules = self.load_schedules()
        self.recheck_all_schedules()

    def remove_schedule(self, idx):
        print(idx)
        if 0 <= idx < len(self.schedules):
            del self.schedules[idx]

        data = {'entry': self.schedules}
        with open('Source/Resources/schedules.txt', 'wb') as file:
            pickle.dump(data, file)
        self.recheck_all_schedules()

    def recheck_all_schedules(self):
        if self.process is not None:
            for process in self.process:
                self.root.after_cancel(process)
        if self.schedules is not None:
            for idx, schedule in enumerate(self.schedules):
                print(schedule)
                self.try_start_alarm(idx, int(schedule["hour"]), int(schedule["minute"]))

    def try_start_alarm(self, idx, hour, minute):
        # todo possible check if processes are cleared, possible code refactoring required
        self.root.after(1000, lambda: self.check_alarm(idx, hour, minute))
        print(self.root.tk.call('after', 'info'))

    def check_alarm(self, idx, hour, minute):
        actual_time = datetime.datetime.now().strftime("%H:%M")
        exact_time = datetime.datetime.now().strftime("%H:%M:%S")
        print(f"current time: {exact_time} still waiting for {hour:02d}:{minute:02d}")
        if actual_time == f"{hour:02d}:{minute:02d}":
            print(f"scheduled time arrived, starting process")
            self.func(idx)
            self.remove_schedule(idx)
            self.cleaner()
            return
        self.process.append(self.root.after(1000, lambda: self.check_alarm(idx, hour, minute)))

    def load_schedules(self, idx=-1):
        if os.path.exists('Source/Resources/schedules.txt'):
            try:
                with open('Source/Resources/schedules.txt', 'r+b') as file:
                    schedules = pickle.load(file)["entry"]
                    if idx != -1 and len(schedules) > idx >= 0:
                        return schedules[idx]
                    else:
                        return schedules
            except EOFError:
                return {}
        else:
            return {}
