# -*- coding: utf-8 -*-
"""
Created on Thu Jun 23 15:02:29 2022

@author: ARMAND Ivan
"""

import pyiast #Fast IAST calculation, published under MIT license
import itertools #to concatenate lists of lists
import numpy as np #dynamic lists
import matplotlib.pyplot as plt #to plot curves
from scipy.optimize import least_squares #to solve non-linear equations
from scipy.optimize import curve_fit
import tkinter as tk #to create a window with frames and canvas
from tkinter import filedialog #to search a file
from tkinter import ttk #to use some widgets
import pandas as pd #to read xlsx and clipboard
import xlsxwriter
from matplotlib.figure import Figure #to draw figures on tkinter frame
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import warnings
warnings.filterwarnings("ignore")
abc=["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"]
all_fitting_methods=["automatic","Langmuir", "Quadratic", "BET", "Henry","TemkinApprox","DSLangmuir","Interpolation"]
def sort_Names(Names):
    """sort a list of the following form : [[index,str] for index in range(n)]"""
    s=[]
    L=[]
    output=[]
    for i in range(len(Names)):
        L.append(Names[i][0])
        s.append(Names[i][0])
    s.sort()
    s=np.array(s)
    L=np.array(L)
    for i in range(len(s)):
        output.append(Names[np.where(L==s[i])[0][0]])
    return output
    
def clear_frame(Frame):
    for child in Frame.winfo_children():
        child.destroy()
        
class Gas:
    """optional Gas(name = ..., number_of_isoT = ..., isoT = [isoT1,isoT2,...], composition=... (between 0 and 1 never 0, never 1)"""
    def __init__(self,**kwargs):
        self.isoT=[] ###contains (T,isoT_array) values
        self.rmse=None
        self.method="automatic"
        for key,value in kwargs.items():
            if key=="name":
                self.name=value
            if key=="number_of_isoT":
                self.number_of_isoT=value
            if key=="isoT":
                self.isoT=value
            if key=="composition":
                self.composition=value
    
    def plot_gas_alone_isoT(self,number_of_points=200,T=293):
        fig=Figure(figsize = (5, 5),dpi = 100)
        fig.clear()
        popup=tk.Tk()
        popup.title("pure isotherm for "+self.get_name()+" at "+str(T)+"°K")
        menubar = tk.Menu(popup)
        popup.config(menu=menubar)
        file_menu = tk.Menu(menubar,tearoff=False)
        file_menu.add_command(label='Exit',command=popup.destroy)
        menubar.add_cascade(label="File",menu=file_menu,underline=0)
        self.set_model_isoT()
        plotx = fig.add_subplot(111)
        isoT=self.get_model_isoT()
        pressure=np.linspace(isoT.df[isoT.pressure_key].min(),isoT.df[isoT.pressure_key].max(),number_of_points)
        plotx.plot(pressure,isoT.loading(pressure))
        plotx.scatter(isoT.df[isoT.pressure_key],isoT.df[isoT.loading_key])
        #for error calculus :
        X=isoT.df[isoT.pressure_key]
        Y=isoT.df[isoT.loading_key]
        Yf=isoT.loading(X)
        Y,Yf=np.array(Y),np.array(Yf)
        nsd=7
        p=1
        rmse_t=number_to_string(RMSE(Y,Yf), nsd)
        R2_t=number_to_string(Coefficient_of_nondetermination(Y, Yf), nsd)
        X2_t=number_to_string(chi_square(Y, Yf), nsd)
        errsq_t=number_to_string(ERRSQ(Y, Yf), nsd)
        easb_t=number_to_string(EASB(Y,Yf), nsd)
        hybrid_t=number_to_string(HYBRID(Y,Yf,p), nsd)
        mpsd_t=number_to_string(MPSD(Y,Yf,p), nsd)
        are_t=number_to_string(ARE(Y,Yf), nsd)
        Values=[rmse_t,R2_t,X2_t,errsq_t,easb_t,hybrid_t,mpsd_t,are_t]
        error_button=tk.Button(master=popup,command=lambda:self.create_fitting_errors_table(Values, p),text="plot error table")
        error_button.pack(side="top")
        plotx.set_title("Pure isotherm of "+self.get_name()+" at "+str(T)+"°K using "+self.get_fitting_method()+" method")
        plotx.set_xlabel("Pressure (bar)")
        plotx.set_ylabel("Loading(mmol/g)")
        canvas=FigureCanvasTkAgg(fig,master=popup)
        canvas.get_tk_widget().pack(side=tk.RIGHT,fill=tk.BOTH,expand=tk.YES)
        canvas.draw()
        popup.mainloop()
        
        
    def create_fitting_errors_table(self,Values,p):
        """give the errors table for a fitting method"""
        popup=tk.Tk()
        popup.title(self.get_fitting_method()+" error table for "+self.get_name())
        Functions=["RMSE","R²","X²","ERRSQ","EASB","HYBRID","MPSD","ARE"]
        tableframe=tk.Frame(popup)
        tableframe.pack(fill = tk.BOTH,expand=tk.YES)
        scrollv = tk.Scrollbar(tableframe,orient="vertical")
        scrollv.pack(side="right", fill="y")
        scrollh = tk.Scrollbar(tableframe,orient='horizontal')
        scrollh.pack(side="bottom", fill="x")
        table=ttk.Treeview(tableframe,yscrollcommand=scrollv.set, xscrollcommand = scrollh.set)
        table.column("#0", width=0,  stretch="no")
        scrollv.config(command=table.yview)
        scrollh.config(command=table.xview)
        table["columns"]=("col1","col2")
        table.column("col1",anchor="center",width=80)
        table.heading("col1",text=self.get_fitting_method()+" error table for "+self.get_name(),anchor="center")
        table.column("col2",anchor="center", width=80)
        table.heading("col2",text="value",anchor="center")
        for f in range(0,len(Functions)):
            val=[Functions[f],Values[f]]
            table.insert(parent='',index='end',iid=f,text='',values=tuple(val))
        table.pack(fill = tk.BOTH,expand=tk.YES)
        menubar = tk.Menu(popup)
        popup.config(menu=menubar)
        file_menu = tk.Menu(menubar,tearoff=False)
        file_menu.add_command(label='Exit',command=popup.destroy)
        menubar.add_cascade(label="File",menu=file_menu,underline=0)
        popup.mainloop()
        
    def set_name(self,name):
        self.name=name
    
    def set_index(self,index):
        """a sort of ID"""
        self.index=index
        
    def set_isoT(self,T,isoT_table):
        if type(isoT_table)!=type(None):
            self.isoT.append((T,isoT_table))
    
    def set_composition(self,comp):
        self.composition=comp
    
    def set_fitting_method(self,method):
        self.method=method
        
    def set_model_isoT(self):
        T,isoT_table=self.isoT[0]
        final_model="Langmuir"
        s=10
        pk="Relative Pressure (p/p°)"
        lk="Quantity Adsorbed (mmol/g)"
        if self.method=="automatic":
            for M in all_fitting_methods[1:-1]:
                try:
                    model_isoT=pyiast.ModelIsotherm(isoT_table,
                                                        loading_key=lk,
                                                        pressure_key=pk,
                                                        model=M)
                    if model_isoT.rmse<s:
                        s=model_isoT.rmse
                        final_model=M
                    self.rmse=s
                except:
                    pass
            self.model_isoT=pyiast.ModelIsotherm(isoT_table,
                                                loading_key=lk,
                                                pressure_key=pk,
                                                model=final_model)
            self.method=final_model
        else:
            if self.method=="Interpolation":
                self.model_isoT=pyiast.InterpolatorIsotherm(isoT_table,
                                                    loading_key=lk,
                                                    pressure_key=pk,
                                                    fill_value=isoT_table[lk].max())
            else:
                self.model_isoT=pyiast.ModelIsotherm(isoT_table,
                                                    loading_key=lk,
                                                    pressure_key=pk,
                                                    model=self.method)
                
    def get_name(self):
        return self.name
    
    def get_index(self):
        return self.index
    
    def get_isoT(self):
        """isoT=[(T,isoT_table)+...]"""
        return self.isoT
    
    def get_composition(self):
        return self.composition
    
    def get_model_isoT(self):
        return self.model_isoT
    
    def get_fitting_method(self):
        return self.method
    
class IAST:
    def __init__(self):
        self.number_of_datas_entered=0
        self.Px=0.001
        self.Py=10
        self.number_iast_points=100
        self.root = tk.Tk()
        self.root.title('IAST calculus')
        self.root.geometry('600x400+50+50')
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        file_menu = tk.Menu(menubar,tearoff=False)
        file_menu.add_command(label='Reset all',command=lambda:[self.root.destroy(),IAST()])
        file_menu.add_separator()
        file_menu.add_command(label='Exit',command=self.root.destroy)
        about_menu=tk.Menu(menubar,tearoff=False)
        about_menu.add_command(label='how it works',command=self.how_it_works)
        menubar.add_cascade(label="File",menu=file_menu,underline=0)
        menubar.add_cascade(label="About the program",menu=about_menu,underline=0)
        self.pan = tk.Frame(self.root)
        self.pan.pack(expand = tk.YES,fill = tk.BOTH)
        self.graph = tk.Canvas(self.pan,bg='gray')
        self.graph.pack(fill = tk.BOTH,expand=tk.YES,side="bottom")
        self.gases=[]
        self.entry_Frame()
        self.root.mainloop()
        
    
    def entry_Frame(self):
        self.entry_Frame=tk.Frame(self.graph)
        number_of_gases = tk.StringVar()
        tk.Label(self.entry_Frame, text = "number of gases").pack(side="left")
        entry_number_of_gases=tk.Entry(self.entry_Frame, textvariable=number_of_gases)
        entry_number_of_gases.insert(0,2)
        entry_number_of_gases.pack(side="left")
        set_number_of_gases = tk.Button(master = self.entry_Frame, command = lambda:[self.set_gas_names(int(number_of_gases.get())),self.entry_Frame.destroy()],text = "set number of gases",borderwidth=3, relief="solid")
        set_number_of_gases.pack(side="right")
        self.entry_Frame.pack()
    
    def set_gas_names(self,number_of_gases=2):
        self.number_of_gases=number_of_gases
        self.set_gas_names=tk.Frame(self.graph)
        self.gases=([Gas() for i in range(number_of_gases)])
        Frames=[]
        Entries=[]
        for i in range(number_of_gases):
            frame_i=tk.Frame(self.set_gas_names)
            Frames.append(frame_i)
            tk.Label(frame_i, text = "name of gas n°"+str(i+1)).pack(side="left")
            name_of_gas_i=tk.StringVar()
            entry_name_of_gas_i=tk.Entry(frame_i, textvariable=name_of_gas_i)
            entry_name_of_gas_i.insert(0,"gas n°"+str(i+1))
            entry_name_of_gas_i.pack(side="left")
            Entries.append(name_of_gas_i)
            set_name_of_gas_i = tk.Button(master = Frames[i],
                                          command=lambda i=i:[self.gases[i].set_name(Entries[i].get()),
                                                              self.gases[i].set_index(i),
                                                              self.Frame_load_gas_isoT(Frames[i],i)],
                                          text = "set name of gas n°"+str(i+1),
                                          borderwidth=3, 
                                          relief="solid")
            set_name_of_gas_i.pack()
            frame_i.pack()
        self.set_gas_names.pack()
    
    def Frame_load_gas_isoT(self,Frame,i):
        name=self.gases[i].get_name()
        clear_frame(Frame)
        tk.Label(Frame, text = "Temperature of "+name+" =").pack(side="left")
        temperature_of_gas=tk.StringVar()
        entry_Temperature_of_gas=tk.Entry(Frame,textvariable=temperature_of_gas)
        entry_Temperature_of_gas.insert(0,293)
        entry_Temperature_of_gas.pack(side="left")
        tk.Label(Frame, text = "°K").pack(side="left")
        load_data_of_gas_i_at_T=tk.Button(master=Frame,command=lambda:self.gases[i].set_isoT(entry_Temperature_of_gas.get(), self.open_file(Frame,i)),text = "load datas from xlsx file",borderwidth=3, relief="solid")
        load_data_of_gas_i_at_T.pack()

    def ask_if_other_isoT(self,Frame,i):
        tk.Label(Frame,text="for "+self.gases[i].get_name()).pack(side="left")
        set_yes=tk.Button(master=Frame,command=lambda:self.Frame_load_gas_isoT(Frame,i),text="add another isotherm")
        set_yes.pack(side="left")
        set_no=tk.Button(master=Frame,command=lambda:[Frame.destroy(),self.datas_are_loaded()],text="done")
        set_no.pack(side="left")

    def datas_are_loaded(self):
        self.number_of_datas_entered+=1
        if self.number_of_datas_entered==self.number_of_gases:
            clear_frame(self.graph)
            self.ask_composition_Frame()
            
    def ask_composition_Frame(self):
        composition_frame=tk.Frame(master=self.graph)
        pressure_range_frame=tk.Frame(master=composition_frame)
        tk.Label(pressure_range_frame,text="Pressure ∈ [").pack(side="left")
        self.px=tk.StringVar()
        self.py=tk.StringVar()
        self.number_points=tk.StringVar()
        px_entry=tk.Entry(pressure_range_frame,textvariable=self.px)
        px_entry.insert(0,0)
        px_entry.pack(side="left")
        tk.Label(pressure_range_frame,text=";").pack(side="left")
        py_entry=tk.Entry(pressure_range_frame,textvariable=self.py)
        py_entry.insert(0,1)
        py_entry.pack(side="left")
        tk.Label(pressure_range_frame,text="] bar !!!, number of points =").pack(side="left")
        number_points_entry=tk.Entry(pressure_range_frame,textvariable=self.number_points)
        number_points_entry.insert(0,100)
        number_points_entry.pack(side="left")
        pressure_range_frame.pack()
        tk.Label(composition_frame,text="composition ratio (between 0 and 1)").pack()
        self.compositions_entries=[]
        self.fitting_methods_entries=[]
        for i in range(len(self.gases)):
            composition_gas_i_frame=tk.Frame(master=composition_frame)
            tk.Label(composition_gas_i_frame,text=self.gases[i].get_name()).pack(side="left")
            composition_of_gas_i=tk.StringVar()
            self.compositions_entries.append(composition_of_gas_i)
            entry_composition_of_gas_i=tk.Entry(composition_gas_i_frame, textvariable=composition_of_gas_i)
            entry_composition_of_gas_i.insert(0,1/self.number_of_gases)
            entry_composition_of_gas_i.pack(side="left")
            calculate_IAST_for_gas_i=tk.Button(master=composition_gas_i_frame,command=lambda i=i:[self.read_composition_of_all_gases(),self.calculate_IAST_for_gas_i(i)],text="plot iast of "+self.gases[i].get_name())
            calculate_IAST_for_gas_i.pack(side="left")
            select_method=ttk.Combobox(composition_gas_i_frame,values=all_fitting_methods)
            select_method.current(0)
            select_method.pack(side="left")
            self.fitting_methods_entries.append(select_method)
            pure_isoT_fitting=tk.Button(master=composition_gas_i_frame,command=lambda i=i:[self.read_composition_of_all_gases(),
            self.gases[i].plot_gas_alone_isoT(self.number_iast_points)],text="plot pure isotherm")
            pure_isoT_fitting.pack()
            composition_gas_i_frame.pack()
        calculate_IAST_for_every_gas=tk.Button(master=composition_frame,command=lambda:[self.read_composition_of_all_gases(),self.calculate_IAST_for_every_gas()],text="plot iast of every gas")
        calculate_IAST_for_every_gas.pack()
        composition_frame.pack()
        self.read_composition_of_all_gases()
    
    def read_composition_of_all_gases(self):
        for i in range(len(self.compositions_entries)):
            self.gases[i].set_composition(float(self.compositions_entries[i].get().replace(',','.')))
            self.gases[i].set_fitting_method(self.fitting_methods_entries[i].get())
    
    def calculate_IAST_for_gas_i(self,i):
        self.Px=float(self.px.get())
        self.Py=float(self.py.get())
        self.number_iast_points=int(self.number_points.get())
        y=[]
        for k in range(len(self.gases)):
            y.append(self.gases[k].get_composition())
            self.gases[k].set_model_isoT()
        y=np.array(y)
        T,isoT_table=self.gases[i].get_isoT()[0]
        X=np.linspace(self.Px,self.Py,self.number_iast_points)
        all_model_isotherm=[self.gases[k].get_model_isoT() for k in range(len(self.gases))]
        Y=[]
        for total_pressure in X:
            q = pyiast.iast(total_pressure * y, all_model_isotherm)
            Y.append(q)
        Yi=[j[i] for j in Y]
        # pyiast.plot_isotherm(all_model_isotherm[i]) #plot the fit of the isotherm with experimental points
        self.plot_gas_i(X,Yi,self.gases[i].get_name())
    
    def calculate_IAST_for_every_gas(self):
        self.Px=float(self.px.get())
        self.Py=float(self.py.get())
        self.number_iast_points=int(self.number_points.get())
        y=[]
        for k in range(len(self.gases)):
            if self.gases[k].get_composition()!=0:
                y.append(self.gases[k].get_composition())
                self.gases[k].set_model_isoT()
        y=np.array(y)
        X=np.linspace(self.Px,self.Py,self.number_iast_points)
        all_model_isotherm=[self.gases[k].get_model_isoT() for k in range(len(self.gases)) if self.gases[k].get_composition()!=0]
        Y=[]
        for total_pressure in X:
            q = pyiast.iast(total_pressure * y, all_model_isotherm)
            Y.append(q)
        self.plot_every_gas(X,Y)
        
    def plot_gas_i(self,X,Y, name_of_gas_i,T=293):
        fig=Figure(figsize = (5, 5),dpi = 100)
        fig.clear()
        popup=tk.Tk()
        menubar = tk.Menu(popup)
        popup.config(menu=menubar)
        file_menu = tk.Menu(menubar,tearoff=False)
        file_menu.add_command(label="save datas (png and xlsx)",command=lambda:self.save_plot(X,Y,name_of_gas_i,T))
        file_menu.add_command(label='Exit',command=popup.destroy)
        menubar.add_cascade(label="File",menu=file_menu,underline=0)
        plotx = fig.add_subplot(111)
        plotx.set_title("Isotherm for "+name_of_gas_i+" at "+str(T)+"°K in the gas mixture")
        plotx.set_xlabel("Pressure (bar)")
        plotx.set_ylabel("Loading(mmol/g)")
        plotx.plot(X,Y)
        canvas=FigureCanvasTkAgg(fig,master=popup)
        canvas.get_tk_widget().pack(side=tk.RIGHT,fill=tk.BOTH,expand=tk.YES)
        canvas.draw()
        popup.mainloop()
        
    def plot_every_gas(self,X,Y,T=293):
        Compositions=[self.gases[i].get_composition() for i in range(len(self.gases))]
        fig=Figure(figsize = (5, 5),dpi = 100)
        fig.clear()
        popup=tk.Tk()
        menubar = tk.Menu(popup)
        popup.config(menu=menubar)
        file_menu = tk.Menu(menubar,tearoff=False)
        file_menu.add_command(label="save datas (png and xlsx)",command=lambda:self.save_plot_every_gas(X,Y,Compositions,T))
        file_menu.add_command(label='Exit',command=popup.destroy)
        menubar.add_cascade(label="File",menu=file_menu,underline=0)
        plotx = fig.add_subplot(111)
        plotx.set_title("Isotherm of every gas at "+str(T)+"°K in the gas mixture")
        plotx.set_xlabel("Pressure (bar)")
        plotx.set_ylabel("Loading(mmol/g)")
        plotx.plot(X,Y)
        plotx.legend([g.get_name() for g in self.gases if g.get_composition()!=0])
        canvas=FigureCanvasTkAgg(fig,master=popup)
        canvas.get_tk_widget().pack(side=tk.RIGHT,fill=tk.BOTH,expand=tk.YES)
        canvas.draw()
        popup.mainloop()
    
    def save_plot(self,X,Y,name_of_gas_i,T=293):
        plt.clf()
        plt.title("Isotherm of "+name_of_gas_i+" at "+str(T)+"°K in the gas mixture")
        plt.xlabel("Pressure (bar)")
        plt.ylabel("Loading(mmol/g)")
        plt.plot(X,Y)
        try:
            foldername=filedialog.askdirectory()
            plt.savefig(foldername+"/isotherm of "+name_of_gas_i+" in the mixture.png")
            df=pd.DataFrame()
            writer=pd.ExcelWriter(foldername+"/isotherm of "+name_of_gas_i+" in the mixture.xlsx",engine="xlsxwriter")
            df.to_excel(writer,sheet_name="Sheet1")
            workbook=writer.book
            worksheet=writer.sheets["Sheet1"]
            X,Y=np.array(X),np.array(Y)
            worksheet.write(0,0,"Pressure(bar)")
            worksheet.write(0,1,"Loading(mmol/g)")
            worksheet.write(0,2,"gas")
            worksheet.write(0,3,"ratio in the mixture (between 0 and 1)")
            for i in range(len(X)):
                worksheet.write(i+1,0,"="+str(X[i]))
                worksheet.write(i+1,1,"="+str(Y[i]))
            for i in range(len(self.gases)):
                worksheet.write(i+1,2,self.gases[i].get_name())
                worksheet.write(i+1,3,self.gases[i].get_composition())
            chart = workbook.add_chart({'type': 'line'})
            chart.add_series({
            'name': "isotherm of "+name_of_gas_i+" in the mixture",
            'categories': '=Sheet1!$A$2:$A$'+str(len(X)+1),
            'values':     '=Sheet1!$B$2:$B$'+str(len(Y)+1),
        })
            chart.set_x_axis({'name': 'Pressure(bar)'})
            chart.set_y_axis({'name': 'Loading(mmol/g)','major_gridlines': {'visible': False}})
            worksheet.insert_chart('F2', chart)
            workbook.close()
        except:
            pass
    
    def save_plot_every_gas(self,X,Y,Compositions,T=293):
        plt.clf()
        plt.title("Isotherm of every gas at "+str(T)+"°K in the gas mixture")
        plt.xlabel("Pressure (bar)")
        plt.ylabel("Loading(mmol/g)")
        plt.plot(X,Y)
        plt.legend([g.get_name() for g in self.gases])
        try:
            foldername=filedialog.askdirectory()
            plt.savefig(foldername+"/isotherm of every gas in the mixture.png")
            df=pd.DataFrame()
            writer=pd.ExcelWriter(foldername+"/isotherm of every gas in the mixture.xlsx",engine="xlsxwriter")
            df.to_excel(writer,sheet_name="Sheet1")
            workbook=writer.book
            worksheet=writer.sheets["Sheet1"]
            X,Y=np.array(X),np.array(Y)
            worksheet.write(0,0,"gas")
            worksheet.write(1,0,"ratio in the mixture (between 0 and 1)")
            for i in range(len(self.gases)):
                worksheet.write(2,2*i+1,"Pressure(bar)")
                worksheet.write(2,2*i+2,"Loading(mmol/g)")
                worksheet.write(0,2*i+1,self.gases[i].get_name())
                worksheet.write(1,2*i+1,Compositions[i])
                for j in range(len(X)):
                    worksheet.write(j+3,2*i+1,"="+str(X[j]))
                    worksheet.write(j+3,2*i+2,"="+str(Y[j][i]))
            chart = workbook.add_chart({'type': 'line'})
            for i in range(len(self.gases)):
                letterA=abc[2*i+1]
                letterB=abc[2*i+2]
                letterA=""
                letterB=""
                k,j=i,i
                while 2*k+1>-1:
                    letterA+=abc[(2*k+1)%26]
                    k-=26
                while 2*j+2>-1:
                    letterB+=abc[(2*k+2)%26]
                    j-=26
                chart.add_series({
                'name': "isotherm of "+self.gases[i].get_name()+" in the mixture",
                'categories': "=Sheet1!$"+letterA+"$4:$"+letterA+"$"+str(len(X)+3),
                'values':     "=Sheet1!$"+letterB+"$4:$"+letterB+"$"+str(len(X)+3),
            })
            chart.set_x_axis({'name': 'Pressure(bar)"'})
            chart.set_y_axis({'name': 'Loading(mmol/g)','major_gridlines': {'visible': False}})
            worksheet.insert_chart('F2', chart)
            workbook.close()
        except:
            pass
    
    def open_file(self,Frame,name):
        """get access to a file path"""
        A=None
        file_path = filedialog.askopenfilename(filetypes = (("excel files", "*.xlsx"),("All files", "*.*")))
        if file_path[-5:]==".xlsx":
            A=pd.read_excel(file_path)
            clear_frame(Frame)
            self.ask_if_other_isoT(Frame,name)
        return A
    
    def how_it_works(self):
        popup=tk.Tk()
        popup.title('IAST calculus about page')
        popup.geometry('600x400+50+50')   
        menubar = tk.Menu(popup)
        popup.config(menu=menubar)
        file_menu = tk.Menu(menubar,tearoff=False)
        file_menu.add_command(label='Exit',command=popup.destroy)
        menubar.add_cascade(label="File",menu=file_menu,underline=0)
        pan=tk.Frame(popup)
        pan.pack(expand=tk.YES,fill=tk.BOTH)
        Line1="Select number of gases in the mixture"
        Line2="for each gas, add a name and load an xlsx file containing two columns, one with 'Relative Pressure (p/p°)' exactly as title and one with 'Quantity Adsorbed (mmol/g)' exactly as title"
        Line3="Select gas temperature (currently it won't change anything)"
        Line4="click on done"
        Line5="when every datas are loaded, just enter each gas ratio. Keep in mind that the sum has to be 1"
        Line6="now you can plot and save your datas, be aware that the file_name can be the same and will erase previous files"
        Line7="29 June 2022, program made by Ivan ARMAND at Ångström laboratory"
        Nums=["1.","2.","3.","4.","5.","6.","date of the build/author"]
        Text=[Line1,Line2,Line3,Line4,Line5,Line6,Line7]
        lent=[len(l) for l in Text]
        m=max(lent)+500
        for line in range(len(Text)):
            while len(Text[line])<=m:
                Text[line]+=" "
            lframe=tk.Frame(pan)
            tk.Label(lframe, text = Nums[line],fg='red',anchor="w").pack(side="left")
            tk.Label(lframe, text = Text[line],anchor="w").pack(side="right")
            lframe.pack()
        popup.mainloop()
        
def RMSE(Observed,Calculated):
    """Root mean square errors"""
    return np.sqrt(np.mean((Observed-Calculated)**2))

def chi_square(Observed,Calculated):
    s=0
    for i in range(len(Observed)):
        s+=((Calculated[i]-Observed[i])**2)/Observed[i]
    return s

def Coefficient_of_nondetermination(Observed,Calculated):
    s1,s2=0,0
    qmobs=np.mean(Observed)
    for i in range(len(Observed)):
        s1+=(Calculated[i]-Observed[i])**2
        s2+=(Calculated[i]-Observed[i])**2+(Calculated[i]-qmobs)**2
    return 1-s1/s2

def ERRSQ(Observed,Calculated):
    """Sum square of errors"""
    s=0
    for i in range(len(Observed)):
        s+=(Observed[i]-Calculated[i])**2
    return s

def EASB(Observed,Calculated):
    """Sum of absolute errors"""
    s=0
    for i in range(len(Observed)):
        s+=abs(Observed[i]-Calculated[i])
    return s

def HYBRID(Observed,Calculated,p):
    """Hybrid fractional error function, p=number of parameters of the fit function"""
    s=0
    for i in range(len(Observed)):
        s+=((Calculated[i]-Observed[i])**2)/Observed[i]
    return 100*s/(len(Observed)-p)

def MPSD(Observed,Calculated,p):
    """Marquardt’s Percent Standard Deviation"""
    s=0
    for i in range(len(Observed)):
        s+=((Calculated[i]-Observed[i])**2)/Observed[i]
    return (s/(len(Observed)-p))**0.5

def ARE(Observed,Calculated):
    """Average Relative Error"""
    s=0
    for i in range(len(Observed)):
        s+=(Calculated[i]-Observed[i])/Observed[i]
    return 100*s/len(Observed)

def number_to_string(value,nsd):
    """convert the number to string with nsd as number of significant digits"""
    return str(("{:."+str(nsd)+"f}").format(round(value,nsd)))


IAST()