import tkinter

class WindowWidgetTest(tkinter.Tk):
    def __init__(self):
        super().__init__()

        self.title('Widget Test')
        self.geometry('400x100')
        self._init_widets()

    def _init_widets(self):
        frame_top = tkinter.Frame(self, width=400, height=200)
        frame_top.pack()

        self.listbox_value = tkinter.StringVar()
        self.listbox_value.set(['this', 'is', 'a', 'pen'])

        """
        複数のアイテムを選択できるようにするには、selectmodeにtkinter.MULTIPLEかtkinter.EXTENDEDを設定する。
        ここでは、とりあえずtkinter.MULTIPLEを設定しておく。

        tkinterの14. The Listbox widgetのselectmodeによれば、selectmodeに設定できる値と値を設定した場合の挙動はこんな感じ。

        tkinter.BROWSE
        1つのアイテムしか選択できない。
        ただ、何か1つのアイテムをクリックしてドラッグすると、別のアイテムに選択が移動する。
        って、何かいいことあるのかなあ?

        tkinter.SINGLE
        1つのアイテムしか選択できない。
        1つのアイテムが選択された状態で、別のアイテムをクリックすると、後からクリックしたアイテムが選択されたアイテムになる。
        tk.BROWSEのように何か1つのアイテムをクリックしてドラッグしても、別のアイテムに選択が移動しない。
        ま、リストボックスのフツーの使い方ですね。

        tkinter.MULTIPLE
        複数のアイテムが選択できる。
        アイテムをクリックするとクリックしたアイテムが選択されたアイテムになる。
        選択されたアイテムをクリックすると選択が解除される。
        複数のアイテムをまとめて選択したい場合、アイテムを延々とクリックしなければならないので、たくさんのアイテムをまとめて選択したい場合は、次のtkinter.EXTENDEDを使った方がいいかも。

        tkinter.EXTENDED
        複数のアイテムが選択できる。
        何もキーを押さないでアイテムをクリックすると、tkinter.SINGLEと同じで、アイテムをクリックするたびに、後からクリックしたアイテムが選択されたアイテムとなる。
        アイテムが1つ以上選択された状態で、Shiftキーを押しながら別のアイテムをクリックすると、選択されたアイテムからShiftキーを押しながらクリックしたアイテムまでが、すべて選択されたアイテムとなるので、たくさんのアイテムをまとめて選択できる。
        Ctrlキーを押しながらアイテムをクリックすると、tkinter.MULTIPLEと同じで、クリックしたアイテムが選択されたアイテムになる。

        上記のとおり、リストボックスで複数のアイテムを選択できるようにする場合は、tkinter.MULTIPLEかtkinter.EXTENDEDを設定すればいいんだけど、どちらにするかは、やっぱりケースバイケースかなあ…
        ただ、同じウィンドウに2つのリストボックスがあって、リストボックスAはtkinter.MULTIPLEで、リスボックスBはtkinter.EXTENDだと、操作に統一感がないので、どちらかに統一したほうがいいかもね～
        """
        self.listbox = tkinter.Listbox(frame_top, listvariable=self.listbox_value, height=3, selectmode=tkinter.MULTIPLE)
        self.listbox.grid(row=0, column=0)

        event_string = '<Button-1>'

        button = tkinter.Button(frame_top, text='Select Item')
        button.bind(event_string, self.get_item)
        button.grid(row=0, column=1)

    def get_item(self, event):
        """
        選択されたリストボックスのアイテムのインデックスをタプルで返す。
        """
        index_item = self.listbox.curselection()

        if not index_item:
            print('not selected')
        else:
            """
            複数のアイテムが選択される場合があるので、選択されたリストボックスのアイテムのインデックスをタプルのループを回し、インデックス毎にアイテムを取得する。
            ここでは、取得しアイテムをリストに保存するようにしておいたけど、勿論、リストに保存しなくてもいい。
            選択されたリストボックスのアイテムをリスト化するのは、なんか、他にもっといい書き方がありそうだけど…(´･ω･`)
            ま、いっか。
            """
            seleted_items = []
            for index in index_item:
                listbox_item = self.listbox.get(index)
                seleted_items.append(listbox_item)
                print(seleted_items)

    def show(self):
        self.mainloop()


win = WindowWidgetTest()
win.show()