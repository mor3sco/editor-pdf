import os
import PyPDF2
import os.path
from tkinter import *
from functools import partial
from tkinter import filedialog
from PyPDF2 import PdfFileReader
from tkinter import ttk, messagebox
from PyPDF2.pdf import PdfFileWriter


class PDF_Editor:
    def __init__(self, root):
        self.window = root
        self.window.geometry("740x480")
        self.window.title('Editor de PDF')

        # Color Options
        self.color_1 = "white"
        self.color_2 = "dodger blue"
        self.color_3 = "black"
        self.color_4 = 'orange red'

        # Font Options
        self.font_1 = "Helvetica"
        self.font_2 = "Times New Roman"
        self.font_3 = "Kokila"

        self.saving_location = ''


        self.menubar = Menu(self.window)

        edit = Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label='Edit', menu=edit)
        edit.add_command(label='Dividir PDF',command=partial(self.SelectPDF, 1))
        edit.add_command(label='Juntar PDF',command=self.Merge_PDFs_Data)
        edit.add_separator()
        edit.add_command(label='Girar PDF',command=partial(self.SelectPDF, 2))

        about = Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label='Sobre', menu=about)
        about.add_command(label='Sobre', command=self.AboutWindow)

        exit = Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label='Fechar', menu=exit)
        exit.add_command(label='Fechar', command=self.Exit)

        self.window.config(menu=self.menubar)

        self.frame_1 = Frame(self.window,bg=self.color_2,width=740,height=480)
        self.frame_1.place(x=0, y=0)
        self.Home_Page()

    def AboutWindow(self):
        messagebox.showinfo("Editor de PDF", \
        "Editor de PDF\nEvandro Moresco\n29/05/2022")

    def ClearScreen(self):
        for widget in self.frame_1.winfo_children():
            widget.destroy()

    def Update_Path_Label(self):
        self.path_label.config(text=self.saving_location)

    def Update_Rotate_Page(self):
        self.saving_location = ''
        self.ClearScreen()
        self.Home_Page()

    def Exit(self):
        self.window.destroy()
   
    

    def Home_Page(self):
        self.ClearScreen()
    
        self.split_button = Button(self.frame_1, text='Dividir',\
        font=(self.font_1, 25, 'bold'), bg="yellow", fg="black", width=8,\
        command=partial(self.SelectPDF, 1))
        self.split_button.place(x=260, y=80)

      
        self.merge_button = Button(self.frame_1, text='Juntar', \
        font=(self.font_1, 25, 'bold'), bg="yellow", fg="black", \
        width=8, command=self.Merge_PDFs_Data)
        self.merge_button.place(x=260, y=160)

        self.rotation_button = Button(self.frame_1, text='Girar', \
        font=(self.font_1, 25, 'bold'), bg="yellow", fg="black", \
        width=8, command=partial(self.SelectPDF, 2))
        self.rotation_button.place(x=260, y=240)


    def SelectPDF(self, to_call):
        self.PDF_path = filedialog.askopenfilename(initialdir = "/", \
        title = "Selecione um arquivo PDF", filetypes = (("Arquivos PDF", "*.pdf*"),))
        if len(self.PDF_path) != 0:
            if to_call == 1:
                self.Split_PDF_Data()
            else:
                self.Rotate_PDFs_Data()

    def SelectPDF_Merge(self):
        self.PDF_path = filedialog.askopenfilenames(initialdir = "/", \
        title = "Selecione um arquivo PDF", filetypes = (("Arquivos PDF", "*.pdf*"),))
        for path in self.PDF_path:
            self.PDF_List.insert((self.PDF_path.index(path)+1), path)

    def Select_Directory(self):

        self.saving_location = filedialog.askdirectory(title = \
        "Selecione um local")
        self.Update_Path_Label()


    def Split_PDF_Data(self):
        pdfReader = PyPDF2.PdfFileReader(self.PDF_path)
        total_pages = pdfReader.numPages

        self.ClearScreen()

        home_btn = Button(self.frame_1, text="Início", \
        font=(self.font_1, 10, 'bold'), command=self.Home_Page)
        home_btn.place(x=10, y=10)


        header = Label(self.frame_1, text="Dividir PDF", \
        font=(self.font_3, 25, "bold"), bg=self.color_2, fg=self.color_1)
        header.place(x=265, y=15)

        
        self.pages_label = Label(self.frame_1, \
        text=f"Número total de páginas: {total_pages}", \
        font=(self.font_2, 20, 'bold'), bg=self.color_2, fg=self.color_3)
        self.pages_label.place(x=40, y=70)

        From = Label(self.frame_1, text="De", \
        font=(self.font_2, 16, 'bold'), bg=self.color_2, fg=self.color_1)
        From.place(x=40, y= 120)

        self.From_Entry = Entry(self.frame_1, font=(self.font_2, 12, 'bold'), \
        width=8)
        self.From_Entry.place(x=40, y= 160)

        # To Label
        To = Label(self.frame_1, text="Para", font=(self.font_2, 16, 'bold'), \
        bg=self.color_2, fg=self.color_1)
        To.place(x=160, y= 120)

        self.To_Entry = Entry(self.frame_1, font=(self.font_2, 12, 'bold'), \
        width=8)
        self.To_Entry.place(x=160, y= 160)

        Cur_Directory = Label(self.frame_1, text="Local de armazenamento", \
        font=(self.font_2, 16, 'bold'), bg=self.color_2, fg=self.color_1)
        Cur_Directory.place(x=300, y= 120)

        self.path_label = Label(self.frame_1, text='/', \
        font=(self.font_2, 16, 'bold'), bg=self.color_2, fg=self.color_3)
        self.path_label.place(x=300, y= 160)

        select_loc_btn = Button(self.frame_1, text="Selecionar local", \
        font=(self.font_1, 8, 'bold'), command=self.Select_Directory)
        select_loc_btn.place(x=320, y=200)

        split_button = Button(self.frame_1, text="Dividir", \
        font=(self.font_3, 16, 'bold'), bg=self.color_4, fg=self.color_1, \
        width=12, command=self.Split_PDF)
        split_button.place(x=250, y=250)
        
    def Merge_PDFs_Data(self):
        self.ClearScreen()
 
        home_btn = Button(self.frame_1, text="Início", \
        font=(self.font_1, 10, 'bold'), command=self.Home_Page)
        home_btn.place(x=10, y=10)


        header = Label(self.frame_1, text="Juntar PDF", \
        font=(self.font_3, 25, "bold"), bg=self.color_2, fg=self.color_1)
        header.place(x=265, y=15)

        select_pdf_label = Label(self.frame_1, text="Selecionar PDF", \
        font=(self.font_2, 20, 'bold'), bg=self.color_2, fg=self.color_3)
        select_pdf_label.place(x=40, y=70)

        open_button = Button(self.frame_1, text="Abrir Pasta", \
        font=(self.font_1, 9, 'bold'), command=self.SelectPDF_Merge)
        open_button.place(x=55, y=110)

        Cur_Directory = Label(self.frame_1, text="Local de armazenamento", \
        font=(self.font_2, 18, 'bold'), bg=self.color_2, fg=self.color_1)
        Cur_Directory.place(x=40, y= 150)


        self.path_label = Label(self.frame_1, text='/', \
        font=(self.font_2, 16, 'bold'), bg=self.color_2, fg=self.color_3)
        self.path_label.place(x=40, y= 190)

        select_loc_btn = Button(self.frame_1, text="Selecionar local", \
        font=(self.font_1, 9, 'bold'), command=self.Select_Directory)
        select_loc_btn.place(x=55, y=225)

        saving_name = Label(self.frame_1, text="Escolha um nome", \
        font=(self.font_2, 18, 'bold'), bg=self.color_2, fg=self.color_1)
        saving_name.place(x=40, y=270)

        self.sv_name_entry = Entry(self.frame_1, \
        font=(self.font_2, 12, 'bold'), width=20)
        self.sv_name_entry.insert(0, 'Resultado')
        self.sv_name_entry.place(x=40, y=310)

        merge_btn = Button(self.frame_1, text="Juntar", \
        font=(self.font_1, 10, 'bold'), command=self.Merge_PDFs)
        merge_btn.place(x=80, y=350)

        listbox_label = Label(self.frame_1, text="Selecionar PDF", \
        font=(self.font_2, 18, 'bold'), bg=self.color_2, fg=self.color_1)
        listbox_label.place(x=482, y=72)

        self.PDF_List = Listbox(self.frame_1,width=40, height=15)
        self.PDF_List.place(x=400, y=110)

        delete_button = Button(self.frame_1, text="Deletar", \
        font=(self.font_1, 9, 'bold'), command=self.Delete_from_ListBox)
        delete_button.place(x=400, y=395)

        more_button = Button(self.frame_1, text="Selecionar Mais", \
        font=(self.font_1, 9, 'bold'), command=self.SelectPDF_Merge)
        more_button.place(x=480, y=395)

    def Rotate_PDFs_Data(self):
        self.ClearScreen()

        pdfReader = PyPDF2.PdfFileReader(self.PDF_path)
        total_pages = pdfReader.numPages

        home_btn = Button(self.frame_1, text="Início", \
        font=(self.font_1, 10, 'bold'), command=self.Home_Page)
        home_btn.place(x=10, y=10)


        header = Label(self.frame_1, text="Girar PDF", \
        font=(self.font_3, 25, "bold"), bg=self.color_2, fg=self.color_1)
        header.place(x=265, y=15)

    
        self.pages_label = Label(self.frame_1, \
        text=f"Número total de páginas: {total_pages}", \
        font=(self.font_2, 20, 'bold'), bg=self.color_2, fg=self.color_3)
        self.pages_label.place(x=40, y=90)

        Cur_Directory = Label(self.frame_1, text="Local de armazenamento", \
        font=(self.font_2, 18, 'bold'), bg=self.color_2, fg=self.color_1)
        Cur_Directory.place(x=40, y= 150)

        self.fix_label = Label(self.frame_1, \
        text="Girar esta página (Número separado por vírgula)", \
        font=(self.font_2, 16, 'bold'), bg=self.color_2, fg=self.color_1)
        self.fix_label.place(x=260, y= 150)

        self.fix_entry = Entry(self.frame_1, \
        font=(self.font_2, 12, 'bold'), width=40)
        self.fix_entry.place(x=260, y=190)


        self.path_label = Label(self.frame_1, text='/', \
        font=(self.font_2, 16, 'bold'), bg=self.color_2, fg=self.color_3)
        self.path_label.place(x=40, y= 190)


        select_loc_btn = Button(self.frame_1, text="Selecionar local", \
        font=(self.font_1, 9, 'bold'), command=self.Select_Directory)
        select_loc_btn.place(x=55, y=225)

        saving_name = Label(self.frame_1, text="Escolha um nome", \
        font=(self.font_2, 18, 'bold'), bg=self.color_2, fg=self.color_1)
        saving_name.place(x=40, y=270)

        self.sv_name_entry = Entry(self.frame_1, \
        font=(self.font_2, 12, 'bold'), width=20)
        self.sv_name_entry.insert(0, 'Resultado')
        self.sv_name_entry.place(x=40, y=310)

        which_side = Label(self.frame_1, text="Rotação", \
        font=(self.font_2, 16, 'bold'), bg=self.color_2, fg=self.color_1)
        which_side.place(x=260, y=230)

        text = StringVar()
        self.alignment = ttk.Combobox(self.frame_1, textvariable=text)
        self.alignment['values'] = ('ClockWise',
                                    'Anti-ClockWise'
                                    )
        self.alignment.place(x=260, y=270)

        rotate_button = Button(self.frame_1, text="Girar", \
        font=(self.font_3, 16, 'bold'), bg=self.color_4, \
        fg=self.color_1, width=12, command=self.Rotate_PDFs)
        rotate_button.place(x=255, y=360)

    def Split_PDF(self):
        if self.From_Entry.get() == "" and self.To_Entry.get() == "":
            messagebox.showwarning("Aviso!", \
            "Mencione o intervalo de páginas\n você quer dividir")
        else:
            from_page = int(self.From_Entry.get()) - 1
            to_page = int(self.To_Entry.get())

            pdfReader = PyPDF2.PdfFileReader(self.PDF_path)

            for page in range(from_page, to_page):
                pdfWriter = PdfFileWriter()
                pdfWriter.addPage(pdfReader.getPage(page))

                splitPage = os.path.join(self.saving_location,f'{page+1}.pdf')
                resultPdf = open(splitPage, 'wb')
                pdfWriter.write(resultPdf)

            resultPdf.close()
            messagebox.showinfo("Sucesso!","O arquivo PDF foi dividido")

            self.saving_location = ''
            self.total_pages = 0
            self.ClearScreen()
            self.Home_Page()


    def Merge_PDFs(self):
        if len(self.PDF_path) == 0:
            messagebox.showerror("Erro!", "Selecione primeiro os PDFs")
        else:
            if self.saving_location == '':
                curDirectory = os.getcwd()
            else:
                curDirectory = str(self.saving_location)

            presentFiles = list()

            for file in os.listdir(curDirectory):
                presentFiles.append(file)
        
            checkFile = f'{self.sv_name_entry.get()}.pdf'

            if checkFile in presentFiles:
                messagebox.showwarning('Aviso!', \
                "Selecione outro nome de arquivo para salvar")
            else:
                pdfWriter = PyPDF2.PdfFileWriter()

                for file in self.PDF_path:
                    pdfReader = PyPDF2.PdfFileReader(file)
                    numPages = pdfReader.numPages
                    for page in range(numPages):
                        pdfWriter.addPage(pdfReader.getPage(page))

                mergePage = os.path.join(self.saving_location, \
                f'{self.sv_name_entry.get()}.pdf')
                mergePdf = open(mergePage, 'wb')
                pdfWriter.write(mergePdf)

                mergePdf.close()
                messagebox.showinfo("Sucesso!", \
                "Os PDFs foram mesclados com sucesso")

                self.saving_location = ''
                self.ClearScreen()
                self.Home_Page()

    def Delete_from_ListBox(self):
        try:
            if len(self.PDF_path) < 1:
                messagebox.showwarning('Aviso!', \
                'Não há mais arquivos para excluir')
            else:
                for item in self.PDF_List.curselection():
                    self.PDF_List.delete(item)
                    
                self.PDF_path = list(self.PDF_path)
                del self.PDF_path[item]
        except Exception:
            messagebox.showwarning('Atenção!',"Por favor, selecione os PDFs primeiro")

    def Rotate_PDFs(self):
        need_to_fix = list()

        if self.fix_entry.get() == "":
            messagebox.showwarning("Aviso!", \
            "Insira o número da página separado por vírgula")
        else:
            for page in self.fix_entry.get().split(','):
                    need_to_fix.append(int(page))

            if self.saving_location == '':
                curDirectory = os.getcwd()
            else:
                curDirectory = str(self.saving_location)

            presentFiles = list()

            for file in os.listdir(curDirectory):
                presentFiles.append(file)

            checkFile = f'{self.sv_name_entry.get()}.pdf'

            if checkFile in presentFiles:
                messagebox.showwarning('Aviso!', \
                "Selecione outro nome de arquivo para salvar")
            else:
                if self.alignment.get() == 'Horário':
                    pdfReader = PdfFileReader(self.PDF_path)
                    pdfWriter = PdfFileWriter()

                    rotatefile = os.path.join(self.saving_location, \
                    f'{self.sv_name_entry.get()}.pdf')
                    fixed_file = open(rotatefile, 'wb')

                    for page in range(pdfReader.getNumPages()):
                        thePage = pdfReader.getPage(page)
                        if (page+1) in need_to_fix:
                            thePage.rotateClockwise(90)

                        pdfWriter.addPage(thePage)

                    pdfWriter.write(fixed_file)
                    fixed_file.close()
                    messagebox.showinfo('Sucesso', 'Giro Completo')
                    self.Update_Rotate_Page()

                elif self.alignment.get() == 'Anti-Horário':
                    pdfReader = PdfFileReader(self.PDF_path)
                    pdfWriter = PdfFileWriter()

                    rotatefile = os.path.join(self.saving_location, \
                    f'{self.sv_name_entry.get()}.pdf')
                    fixed_file = open(rotatefile, 'wb')

                    for page in range(pdfReader.getNumPages()):
                        thePage = pdfReader.getPage(page)
                        if (page+1) in need_to_fix:
                            thePage.rotateCounterClockwise(90)

                        pdfWriter.addPage(thePage)

                    pdfWriter.write(fixed_file)
                    fixed_file.close()
                    messagebox.showinfo('Sucesso','Giro completo')
                    self.Update_Rotate_Page()
                else:
                    messagebox.showwarning('Aviso!', \
                    "Please Select a Right Alignment")

# The main function
if __name__ == "__main__":
    root = Tk()
    # Creating a CountDown class object
    obj = PDF_Editor(root)
    root.mainloop()
