import ttkbootstrap as ttk

class StringListVar:
    def __init__(self, master=None, values=None, callback=None):
        self.master = master
        self.values = values or []
        self.callback = callback
        self.trace_variable = ttk.StringVar(master)
        self.trace_variable.trace_add("write", self._update_values)
        self.trace_variable.trace_add("read", self._update_values)

    def _update_values(self, *args):
        self.values = self.trace_variable.get().split(',')
        if self.callback:
            self.callback()

    def set(self, values):
        self.values = values
        self.trace_variable.set(','.join(values))

    def get(self):
        return self.values