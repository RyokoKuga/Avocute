import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import configparser
import subprocess
import webbrowser
import shutil
import glob
import os
import csv

######################
#####   Tools   ##### 

# BatchLauncher #
# サブウインドウ定義
def create_BL_Window():
    #####   関数(BatchLauncher)   #####
    # datファイル保存＆登録
    def bl_save_Input(bl_data):
        global bl_dat_path
        # ファイル選択ダイアログ表示
        typ = [('MOPAC','*.dat')] 
        dir = 'C:\\pg'
        bl_dat_path = tk.filedialog.asksaveasfilename(filetypes = typ, initialdir = dir, defaultextension = "dat")

        if len(bl_dat_path) != 0:
            # ファイルが選択された場合データを書き込み
            bl_f = open(bl_dat_path, "w")
            bl_f.write(bl_data)
            bl_f.close()
            # パスをツリービューに登録
            bl_entry_table.insert("", "end", values = ("Pending", bl_dat_path))
            # ステータス更新
            bl_statusbar["text"] = "  Let's Start!!"
        else:
            pass   

    # テキストボックス(dat作成欄)からテキスト取得
    def bl_write_button_func():
        # テキストボックスから文字列取得
        bl_data = txtbox.get("1.0", "end")
        # AUXチェック
        if mode_chkvalue.get() and "AUX" in bl_data:
            bl_data = bl_data.replace("AUX", "")
            bl_save_Input(bl_data)
        else:
            bl_save_Input(bl_data)
            
    # バッチ計算後処理
    def post_processing():
        # ツリービュー更新
        bl_entry_table.delete(child)
        bl_entry_table.insert("", "end", values = ("Done", bl_path))

        # 終了後に一時ファイル削除
        bl_file_list = glob.glob("FOR0*")
        for file in bl_file_list:
            os.remove(file)
    
    # FOR023をmgfにリネーム,移動
    def bl_FOR023_rm():
        # FOR023をmgfにリネーム
        if(os.path.exists("FOR023")):
            bl_mgf_name = os.path.splitext(os.path.basename(bl_path))[0] + ".mgf"
            os.rename("FOR023", bl_mgf_name)
            # arcファイル移動
            bl_mgf_move_path = os.path.dirname(bl_path)
            bl_mgf_path = os.path.join(bl_mgf_move_path, bl_mgf_name)
            shutil.move(bl_mgf_name, bl_mgf_path)

    # FOR012をarcにリネーム,移動
    def bl_FOR012_rm():
        # FOR012をarcにリネーム
        if(os.path.exists("FOR012")):
            bl_arc_name = os.path.splitext(os.path.basename(bl_path))[0] + ".arc"
            os.rename("FOR012", bl_arc_name)
            # arcファイル移動
            bl_arc_move_path = os.path.dirname(bl_path)
            bl_arc_path = os.path.join(bl_arc_move_path, bl_arc_name)
            shutil.move(bl_arc_name, bl_arc_path)

    # FOR006をoutにリネーム,移動
    def bl_FOR006_rm():
        # FOR006をarcにリネーム
        if(os.path.exists("FOR006")):
            bl_out_name = os.path.splitext(os.path.basename(bl_path))[0] + ".out"
            os.rename("FOR006", bl_out_name)
            # outファイル移動
            bl_out_move_path = os.path.dirname(bl_path)
            bl_out_path = os.path.join(bl_out_move_path, bl_out_name)
            shutil.move(bl_out_name, bl_out_path)
            
    # バッチ計算実行
    def batch_calc():
        global bl_path, child
        
        for child in  bl_entry_table.get_children():
            bl_path = bl_entry_table.item(child)["values"][1]
            # FOR005へdat内容のコピー
            shutil.copy(bl_path, "FOR005")
            subprocess.run(["mopac606.exe", "FOR005"])
            
            # 終了後にFOR006をoutにリネーム,移動
            bl_FOR006_rm()
            # 終了後にFOR012をarcにリネーム,移動
            bl_FOR012_rm()
            # 終了後にFOR023をmgfにリネーム,移動
            bl_FOR023_rm()

            # バッチ計算後処理
            post_processing()
            
        # 終了後にステータス更新
        bl_statusbar["text"] = "  The entire batch process is complete!!"
            
    # ステータスチェック
    def bl_status_chk():
        # ステータスから文字列取得
        bl_status = bl_statusbar.cget("text")
        if bl_status == "  Let's Start!!":
            # ステータス更新
            bl_statusbar["text"] = "  wait..."
            bl_window.update_idletasks()
            
            # MOPACバッチ実行
            batch_calc()
        
    # MOPACでdatファイルを実行
    def bl_mopac_run():
        # MOPAC(exe)が存在するかチェック
        if os.path.exists("mopac606.exe"):
            # 存在する場合はステータスチェックから実行
            bl_status_chk()
        
        # MOPACが見つからないときはエラー表示
        else:
            messagebox.showerror('Error', 'No mopac.exe!!')

    # ツリービューの値を全て削除
    def path_delete():
        for child in  bl_entry_table.get_children():
            bl_entry_table.delete(child)
        # 終了後にステータス更新
        bl_statusbar["text"] = "  No input file!!"

        
    #####   GUI(BatchLauncher)   #####
    # ウインドウ作成
    bl_window = tk.Toplevel(root)
    bl_window.geometry("650x370")
    bl_window.resizable(width = False, height = False)
    
    # フレーム作成
    bl_entry_frame = ttk.Frame(bl_window, paddin = 5)
    bl_entry_frame.pack(padx = 5)
    
    # ラベル作成
    blname_label = tk.Label(bl_entry_frame, text = "Batch Launcher:").pack(padx = 5, pady=5, anchor = tk.W)
    
    # dat登録用ツリービュー作成
    bl_entry_table = ttk.Treeview(bl_entry_frame)
    # 列の作成
    bl_entry_table["column"] = (1, 2)
    bl_entry_table["show"] = "headings"
    # ヘッダーテキスト
    bl_entry_table.heading(1, anchor="w", text = "Status")
    bl_entry_table.heading(2, anchor="w", text = "Name")
    # 列幅
    bl_entry_table.column(1, width = 90)
    bl_entry_table.column(2, width = 500)
    
    # スクロールバー作成
    bl_yscroll = tk.Scrollbar(bl_entry_frame, orient = tk.VERTICAL, command = bl_entry_table.yview)
    bl_yscroll.pack(side = tk.RIGHT, fill = "y")
    bl_entry_table["yscrollcommand"] = bl_yscroll.set
    
    # ツリービュー設置
    bl_entry_table.pack(fill = tk.BOTH)
    
    # ステータスバー作成
    bl_statusbar = tk.Label(bl_window, bd = 1, text =  "  No input file!!", relief = tk.SUNKEN, anchor = tk.W)
    bl_statusbar.pack(side = tk.BOTTOM, fill=tk.X)
   
    # Runボタン作成
    global blrun_image
    blrun_image = tk.PhotoImage(file = "Images/Run.png")
    bl_run_button = tk.Button(bl_window, text=" Run", command = bl_mopac_run,
                             image = blrun_image, compound = tk.LEFT,  width = 107, height = 30)
    bl_run_button.pack(side = tk.RIGHT, padx=(10, 38), pady=(0, 15))

    # Addボタン作成
    global bladd_image
    bladd_image = tk.PhotoImage(file = "Images/File.png")
    add_button = tk.Button(bl_window, text=" Add", command = bl_write_button_func,
                          image = bladd_image, compound = tk.LEFT,  width = 107, height = 30)
    add_button.pack(side = tk.RIGHT, padx=(0, 0), pady=(0, 15))

    # Resetボタン作成
    global blreset_image
    blreset_image = tk.PhotoImage(file = "Images/Reset.png")
    reset_button = tk.Button(bl_window, text=" Reset", command = path_delete,
                            image = blreset_image, compound = tk.LEFT,  width = 107, height = 30)
    reset_button.pack(side = tk.LEFT, padx=(20, 0), pady=(0, 15))


# AutoTsSearcher #
# サブウインドウ定義
def create_ATS_Window():

    ######   関数(AutoTsSearcher)  #####
    # reactant.dat保存
    def r_save_Input(ats_data): 
        global r_dat_path

        # ファイル選択ダイアログの表示    
        typ = [('MOPAC','*.dat')]
        dir = 'C:\\pg'
        # reactant.datのパス
        r_dat_path = filedialog.asksaveasfilename(filetypes = typ, initialdir = dir, defaultextension = "dat")

        if len(r_dat_path) != 0:
            # ファイルが選択された場合データを書き込み
            rf = open(r_dat_path, "w")
            rf.write(ats_data)
            rf.close()
            # 既存のパスを削除
            r_path_box.delete(0, tk.END)
            # パスをエントリー中に表示
            r_path_box.insert(tk.END, r_dat_path)
        else:
            pass

    # product.dat保存
    def p_save_Input(ats_data): 
        global p_dat_path

        # ファイル選択ダイアログの表示    
        typ = [('MOPAC','*.dat')]
        dir = 'C:\\pg'
        # product.datのパス
        p_dat_path = filedialog.asksaveasfilename(filetypes = typ, initialdir = dir, defaultextension = "dat")

        if len(p_dat_path) != 0:
            # ファイルが選択された場合データを書き込み
            pf = open(p_dat_path, "w")
            pf.write(ats_data)
            pf.close()
            # 既存のパスを削除
            p_path_box.delete(0, tk.END)
            # パスをエントリー中に表示
            p_path_box.insert(tk.END, p_dat_path)
            # ステータス更新
            ats_statusbar["text"] = "  Let's Start!!"
        else:
            pass         

    # テキストボックス(dat作成欄)からテキストを取得
    def r_write_button_func():
        # テキストボックスから文字列取得
        ats_data = txtbox.get("1.0", "end")
        # AUXチェック
        if mode_chkvalue.get() and "AUX" in ats_data:
            ats_data = ats_data.replace("AUX", "")
            r_save_Input(ats_data)
        else:
            r_save_Input(ats_data)        

    def p_write_button_func():
        # テキストボックスから文字列取得
        ats_data = txtbox.get("1.0", "end")
        # AUXチェック
        if mode_chkvalue.get() and "AUX" in ats_data:
            ats_data = ats_data.replace("AUX", "")
            p_save_Input(ats_data)
        else:
            p_save_Input(ats_data)

    # テキストボックス(reactant,productのパス)のリセット
    def ats_path_delete():
        r_path_box.delete(0, tk.END)
        p_path_box.delete(0, tk.END)
        # ステータスの更新
        ats_statusbar["text"] = "  No input file!!"

    # reactant,product(dat)パス取得
    def get_rp_path():
        global r_path, p_path
        # テキストボックスからパス取得
        r_path = r_path_box.get()
        p_path = p_path_box.get()

    # 一時ファイル削除
    def tmpfile_delete():
        ats_file_list = glob.glob("FOR0*")
        for ats_file in ats_file_list:
            os.remove(ats_file)

    # reactant.arcの移動
    def r_arc_move():
        global r_move_path

        r_move_path = os.path.dirname(r_dat_path)
        r_arc_path = os.path.join(r_move_path, r_arc_name)
        shutil.move(r_arc_name, r_arc_path)

    # product.arcの移動
    def p_arc_move():
        p_move_path = os.path.dirname(p_dat_path)
        p_arc_path = os.path.join(p_move_path, p_arc_name)
        shutil.move(p_arc_name, p_arc_path)

    # saddle.arcの移動
    def s_arc_move():
        s_arc_path = os.path.join(r_move_path, s_arc_name)
        shutil.move(s_arc_name, s_arc_path)

    # ts.arcの移動
    def ts_arc_move():
        ts_arc_path = os.path.join(r_move_path, ts_arc_name)
        shutil.move(ts_arc_name, ts_arc_path)

    # saddle(dat,out)の移動
    def s_datout_move():
        global s_out_path
        
        s_dat_path = os.path.join(r_move_path, s_filename)
        shutil.move(s_filename, s_dat_path)
        s_out_path = os.path.splitext(r_path)[0] +"_saddle.out"
        shutil.move(os.path.basename(s_out_path), s_out_path)

    # ts(dat,out)の移動
    def ts_datout_move():
        ts_dat_path = os.path.join(r_move_path, ts_filename)
        shutil.move(ts_filename, ts_dat_path)
        ts_out_path = os.path.splitext(r_path)[0] +"_ts.out"
        shutil.move(os.path.basename(ts_out_path), ts_out_path)

    # frq(dat,out)の移動
    def frq_datout_move():
        global frq_out_path
        
        frq_dat_path = os.path.join(r_move_path, frq_filename)
        shutil.move(frq_filename, frq_dat_path)
        frq_out_path = os.path.splitext(r_path)[0] +"_frq.out"
        shutil.move(os.path.basename(frq_out_path), frq_out_path)

    # outputエラー確認(frq)
    def frq_stts_chk():
        # outputの内容をリスト化
        with open(frq_out_path,"r") as frq_outfile_stts:
            frq_stts_list = frq_outfile_stts.readlines()
        # outputからエラーが無いか確認
        if " == MOPAC DONE ==\n" in frq_stts_list:
            # ステータス更新
            ats_statusbar["text"] = "  MOPAC terminated!! == MOPAC DONE =="
        else:
            # ステータス更新
            ats_statusbar["text"] = "  MOPAC terminated!!  Error: Check the output!!" 

    # frq実行
    def frq_run():
        # FOR005へdatの内容をコピー
        shutil.copy(frq_filename, "FOR005")
        # MOPAC実行(frq)
        subprocess.run(["mopac606.exe", "FOR005"])
        
        # FOR006をoutにリネーム
        if(os.path.exists("FOR006")):
            frq_out_name = os.path.splitext(os.path.basename(r_path))[0] + "_frq.out"
            os.rename("FOR006", frq_out_name)

        # arcファイル移動
        r_arc_move()
        p_arc_move()
        s_arc_move()
        ts_arc_move()

        # saddle,ts,frq(dat,out)の移動
        s_datout_move()
        ts_datout_move()
        frq_datout_move()

        # 一時ファイル削除
        tmpfile_delete()
        # 終了後にステータス更新
        frq_stts_chk()

    # frq.dat作成
    def frq_gen():
        global frq_filename

        # frqの書込み先指定
        frq_filename = os.path.splitext(os.path.basename(r_path))[0] +"_frq.dat"
        frq_file = open(frq_filename,"w+")

        # 出力ファイルの内容をリスト化
        with open(ts_arc_name,"r") as outfile_frq:
            frq_list = outfile_frq.readlines()
        # リストから検索文字列の行番号取得
        for linenum_frq, line_frq in enumerate(frq_list):
            if "FINAL" in line_frq:
                finalnum_frq = linenum_frq
        # 検索文字列の1行目から最終行までリスト内要素を抽出
        keyword_coord__frq = frq_list[(finalnum_frq +1):]
        # 検索文字を置換後ファイルへ書込み
        for frq_line in [frq.replace("TS", "FORCE") for frq in keyword_coord__frq]:
            frq_file.write(frq_line)
        # ファイルを閉じる
        frq_file.close()

        # frq実行
        frq_run()

    # ts.dat実行
    def ts_run():
        global ts_arc_name
        
        # FOR005へdatの内容をコピー
        shutil.copy(ts_filename, "FOR005")
        # MOPAC実行(TS)
        subprocess.run(["mopac606.exe", "FOR005"])
        
        # FOR006をoutにリネーム
        if(os.path.exists("FOR006")):
            ts_out_name = os.path.splitext(os.path.basename(r_path))[0] + "_ts.out"
            os.rename("FOR006", ts_out_name)

        # 計算終了後FOR012をarcにリネーム
        if(os.path.exists("FOR012")):
            ts_arc_name = os.path.splitext(os.path.basename(r_path))[0] + "_ts.arc"
            os.rename("FOR012", ts_arc_name)

            # frq.dat作成
            frq_gen()

        # tsのarcが生成されなかった場合
        else:
            #arcファイルの移動
            r_arc_move()
            p_arc_move()
            s_arc_move()

            # saddle, ts(dat,out)の移動
            s_datout_move()
            ts_datout_move()

            # 終了後に一時ファイル削除
            tmpfile_delete()
            # ステータス更新
            ats_statusbar["text"] = "  Terminated: ts.dat... Check the output!!" 

    # ts.dat作成
    def ts_gen():
        global ts_filename

        # tsの書込み先指定
        ts_filename = os.path.splitext(os.path.basename(r_path))[0] +"_ts.dat"
        ts_file = open(ts_filename,"w+")

        # 出力ファイルの内容をリスト化
        with open(s_arc_name,"r") as outfile_ts:
            ts_list = outfile_ts.readlines()
        # リストから検索文字列の行番号取得
        for linenum_ts, line_ts in enumerate(ts_list):
            if "FINAL" in line_ts:
                finalnum_ts = linenum_ts
        # 検索文字列の1行目から最終行までリスト内要素を抽出
        keyword_coord__ts = ts_list[(finalnum_ts +1):]
        # 検索文字を置換後ファイルへ書込み
        for ts_line in [ts.replace("SADDLE XYZ", "TS") for ts in keyword_coord__ts]:
            ts_file.write(ts_line)
        # ファイルを閉じる
        ts_file.close()

        # ts.dat実行
        ts_run()

    # outputエラー確認(saddle)
    def s_stts_chk():
        # outputの内容をリスト化
        with open(s_out_path,"r") as s_outfile_stts:
            s_stts_list = s_outfile_stts.readlines()
        # outputからエラーが無いか確認
        if " == MOPAC DONE ==\n" in s_stts_list:
            # ステータス更新
            ats_statusbar["text"] = "  MOPAC terminated!! == MOPAC DONE =="
        else:
            # ステータス更新
            ats_statusbar["text"] = "  MOPAC terminated!!  Error: Check the output!!" 

    # ts＆frq実行
    def ts_frq_run():
        # チェックがある場合ts＆frq実行
        if tq_chkvalue.get():
            ts_gen()

        # チェックがない場合終了
        else:
            # arcファイル移動
            r_arc_move()
            p_arc_move()
            s_arc_move()

            # saddle(dat,out)の移動
            s_datout_move()

            # 終了後に一時ファイル削除
            tmpfile_delete()
            # 終了後にステータス更新
            s_stts_chk()

    # saddle.dat実行
    def saddle_run():
        global s_arc_name
        
        # FOR005へdatの内容をコピー
        shutil.copy(s_filename, "FOR005")
        # MOPAC実行(SADDLE)
        subprocess.run(["mopac606.exe", "FOR005"])
        
        # FOR006をoutにリネーム
        if(os.path.exists("FOR006")):
            s_out_name = os.path.splitext(os.path.basename(r_path))[0] + "_saddle.out"
            os.rename("FOR006", s_out_name)

        # 計算終了後FOR012をarcにリネーム
        if(os.path.exists("FOR012")):
            s_arc_name = os.path.splitext(os.path.basename(r_path))[0] + "_saddle.arc"
            os.rename("FOR012", s_arc_name)

            # ts＆frq実行チェック
            ts_frq_run()

        # saddleのarcが生成されなかった場合
        else:
            #arcファイルの移動
            r_arc_move()
            p_arc_move()

            # saddle(dat,out)の移動
            s_datout_move()

            # 終了後に一時ファイル削除
            tmpfile_delete()
            # ステータス更新
            ats_statusbar["text"] = "  Terminated: saddle.dat... Check the output!!"

    # saddle.da作成
    def saddle_gen():
        global s_filename

        # saddlの書込み先指定(reactant)
        s_filename = os.path.splitext(os.path.basename(r_path))[0] +"_saddle.dat"
        saddle_file = open(s_filename,"w+")

        # 出力ファイルの内容をリスト化
        with open(r_arc_name,"r") as outfile_r:
            r_list = outfile_r.readlines()
        # リストから検索文字列の行番号取得
        for linenum_r, line_r in enumerate(r_list):
            if "FINAL" in line_r:
                finalnum_r = linenum_r
        # 検索文字列の1行目から最終行までリスト内要素を抽出
        keyword_coord__r = r_list[(finalnum_r +1):]
        # 検索文字を置換後ファイルへ書込み
        for r_line in [s.replace("EF", "SADDLE XYZ") for s in keyword_coord__r]:
            saddle_file.write(r_line)
        # ファイルを閉じる
        saddle_file.close()

        # saddlの書込み先指定(product)
        saddle_file = open(s_filename,"a+")
        # 出力ファイルの内容をリスト化
        with open(p_arc_name,"r") as outfile_p:
            p_list = outfile_p.readlines()
        # リストから検索文字列の行番号取得
        for linenum_p, line_p in enumerate(p_list):
            if "FINAL" in line_p:
                finalnum_p = linenum_p
        # 検索文字列の4行目から最終行までリスト内要素を抽出
        coord__p = p_list[(finalnum_p +4):]
        # ファイルへ書込み
        for p_line in coord__p:
            saddle_file.write(p_line)
        # ファイルを閉じる
        saddle_file.close()

        # saddle.dat実行
        saddle_run()

    # product.dat実行
    def product_run():
        global p_arc_name

        # FOR005へdatの内容をコピー
        shutil.copy(p_path, "FOR005")
        # MOPAC実行(生成物)
        subprocess.run(["mopac606.exe", "FOR005"])

        # FOR006をoutにリネーム
        if(os.path.exists("FOR006")):
            p_out_name = os.path.splitext(os.path.basename(p_path))[0] + ".out"
            os.rename("FOR006", p_out_name)
            # outファイル移動
            p_out_move_path = os.path.dirname(p_path)
            p_out_path = os.path.join(p_out_move_path, p_out_name)
            shutil.move(p_out_name, p_out_path)

        # 計算終了後FOR012をarcにリネーム
        if(os.path.exists("FOR012")):
            p_arc_name = os.path.splitext(os.path.basename(p_path))[0] + ".arc"
            os.rename("FOR012", p_arc_name)

            # saddle.dat作成
            saddle_gen()

        # 生成物のarcが生成されなかった場合
        else:
            # arcファイル移動
            r_arc_move()

            # 終了後に一時ファイル削除
            tmpfile_delete()
            # ステータス更新
            ats_statusbar["text"] = "  Terminated: Product.dat... Check the output!!"

    # reactant.dat実行
    def reactant_run():
        global r_arc_name

        # ステータスの更新
        ats_statusbar["text"] = "  wait..."
        ats_window.update_idletasks()

        # FOR005へdatの内容をコピー
        shutil.copy(r_path, "FOR005")
        # MOPAC実行(反応物)
        subprocess.run(["mopac606.exe", "FOR005"]) 

        # FOR006をoutにリネーム
        if(os.path.exists("FOR006")):
            r_out_name = os.path.splitext(os.path.basename(r_path))[0] + ".out"
            os.rename("FOR006", r_out_name)
            # outファイル移動
            r_out_move_path = os.path.dirname(r_path)
            r_out_path = os.path.join(r_out_move_path, r_out_name)
            shutil.move(r_out_name, r_out_path)

        # 計算終了後FOR012をarcにリネーム
        if(os.path.exists("FOR012")):
            r_arc_name = os.path.splitext(os.path.basename(r_path))[0] + ".arc"
            os.rename("FOR012", r_arc_name)

            # product.dat実行
            product_run()

        # 反応物のarcが生成されなかった場合
        else:
            # ステータス更新
            ats_statusbar["text"] = "  Terminated: Reactant.dat... Check the output!!"


    # MOPAC実行(ats)
    def ats_mopac_run():
        # MOPAC(exe)が存在するかチェック
        if os.path.exists("mopac606.exe"):
            get_rp_path()
            # パス欄に値がない場合は実行しない
            if len(r_path_box.get()) == 0 or len(p_path_box.get()) == 0:
                messagebox.showerror('Error', 'No Reactant.dat or Product.dat!!')
            else:
                reactant_run()
                
        # MOPACが見つからないときはエラー表示
        else:
            messagebox.showerror('Error', 'No mopac.exe!!')


    #####   GUI(AutoTsSearcher)   #####
    ats_window = tk.Toplevel(root)
    ats_window.geometry("650x240")
    ats_window.resizable(width = False, height = False)

    # ラベル作成
    atsname_label = tk.Label(ats_window, text = "Auto TS Searcher:").place(x = 25, y = 10)
    r_label = tk.Label(ats_window, text = "Reactant:").place(x = 25, y = 40)
    p_label = tk.Label(ats_window, text = "Product:").place(x = 25, y = 70)
    # パス入力欄作成
    r_path_box = tk.Entry(ats_window, width = 65)
    r_path_box.place(x = 95, y = 41)
    p_path_box = tk.Entry(ats_window, width = 65)
    p_path_box.place(x = 95, y = 71)

    # Addボタン
    global folder_image
    folder_image = tk.PhotoImage(file = "Images/File.png")
    r_add_button = tk.Button(ats_window, text="", command = r_write_button_func,
                            image = folder_image, width = 30, height = 20)
    r_add_button.place(x = 563, y = 37)
    p_add_button = tk.Button(ats_window, text="", command = p_write_button_func,
                            image = folder_image, width = 30, height = 20)
    p_add_button.place(x = 563, y = 67)
    # Runボタン
    global atsrun_image
    atsrun_image = tk.PhotoImage(file = "Images/Run.png")
    ats_run_button = tk.Button(ats_window, text = " Run", command = ats_mopac_run,
                               image = atsrun_image, compound = tk.LEFT, width = 540, height = 30)
    ats_run_button.place(x = 50, y = 160)  
    # チェクボタン
    tq_chkvalue = tk.BooleanVar()
    tq_chkvalue.set(True)
    tq_chk = ttk.Checkbutton(ats_window, text = "Transition State ＆ Frequencies", variable = tq_chkvalue)
    tq_chk.place(x = 25, y = 112)
    # Resetボタン
    global ats_reset_image
    ats_reset_image = tk.PhotoImage(file = "Images/Reset.png")
    ats_reset_button = tk.Button(ats_window, text = " Reset", command = ats_path_delete,
                                 image = ats_reset_image, compound = tk.LEFT, width = 107, height = 30)
    ats_reset_button.place(x = 484, y = 105)

    # ステータスバー
    ats_statusbar = tk.Label(ats_window, bd = 1, text =  "  No input file!!", relief = tk.SUNKEN, anchor = tk.W)
    ats_statusbar.pack(side = tk.BOTTOM, fill=tk.X)

    
# CustomKeywords #
# サブウインドウ定義
def create_CK_Window():
    
    # 設定ファイル用のインスタンス
    config = configparser.ConfigParser()
    # 読み込むiniファイルを指定
    config.read("settings.ini")
    
    #####   関数(CustomKeywords)   #####
    # キーワード保存
    def path_save(reg_keyword): 

        if radioValue.get() == 1:
            rdio_1["text"] = reg_keyword

            config["Custom Keywords1"] = {"Keywords1" : reg_keyword}
            # iniファイルへ保存
            with open("settings.ini", "w") as file:
                config.write(file)

        elif radioValue.get() == 2:
            rdio_2["text"] = reg_keyword

            config["Custom Keywords2"] = {"Keywords2" : reg_keyword}
            # iniファイルへ保存
            with open("settings.ini", "w") as file:
                config.write(file)

        elif radioValue.get() == 3:
            rdio_3["text"] = reg_keyword

            config["Custom Keywords3"] = {"Keywords3" : reg_keyword}
            # iniファイルへ保存
            with open("settings.ini", "w") as file:
                config.write(file)    

        elif radioValue.get() == 4:
            rdio_4["text"] = reg_keyword

            config["Custom Keywords4"] = {"Keywords4" : reg_keyword}
            # iniファイルへ保存
            with open("settings.ini", "w") as file:
                config.write(file)

        elif radioValue.get() == 5:
            rdio_5["text"] = reg_keyword

            config["Custom Keywords5"] = {"Keywords5" : reg_keyword}
            # iniファイルへ保存
            with open("settings.ini", "w") as file:
                config.write(file)

        elif radioValue.get() == 6:
            rdio_6["text"] = reg_keyword

            config["Custom Keywords6"] = {"Keywords6" : reg_keyword}
            # iniファイルへ保存
            with open("settings.ini", "w") as file:
                config.write(file)

        elif radioValue.get() == 7:
            rdio_7["text"] = reg_keyword

            config["Custom Keywords7"] = {"Keywords7" : reg_keyword}
            # iniファイルへ保存
            with open("settings.ini", "w") as file:
                config.write(file)

        elif radioValue.get() == 8:
            rdio_8["text"] = reg_keyword

            config["Custom Keywords8"] = {"Keywords8" : reg_keyword}
            # iniファイルへ保存
            with open("settings.ini", "w") as file:
                config.write(file)

        elif radioValue.get() == 9:
            rdio_9["text"] = reg_keyword

            config["Custom Keywords9"] = {"Keywords9" : reg_keyword}
            # iniファイルへ保存
            with open("settings.ini", "w") as file:
                config.write(file)

        elif radioValue.get() == 10:
            rdio_10["text"] = reg_keyword

            config["Custom Keywords10"] = {"Keywords10" : reg_keyword}
            # iniファイルへ保存
            with open("settings.ini", "w") as file:
                config.write(file)
                
    # キーワード取得
    def reg_keyword_func():
        # テキストボックスから文字列取得
        reg_keyword = reg_box.get()

        # 保存関数へ
        path_save(reg_keyword)
    
    # キーワード設定値をテキストボックスへ印字
    def ckto_kw_set():
        # ステータス更新
        statusbar["text"] = "  Let's Start!!"
        # 既存の値削除
        txtbox.delete("1.0","1.end")
        # 値の挿入
        if radioValue.get() == 1:
            txtbox.insert("1.end", rdio_1["text"])
            
        elif radioValue.get() == 2:
            txtbox.insert("1.end", rdio_2["text"])

        elif radioValue.get() == 3:
            txtbox.insert("1.end", rdio_3["text"])

        elif radioValue.get() == 4:
            txtbox.insert("1.end", rdio_4["text"])

        elif radioValue.get() == 5:
            txtbox.insert("1.end", rdio_5["text"])

        elif radioValue.get() == 6:
            txtbox.insert("1.end", rdio_6["text"])

        elif radioValue.get() == 7:
            txtbox.insert("1.end", rdio_7["text"])

        elif radioValue.get() == 8:
            txtbox.insert("1.end", rdio_8["text"])

        elif radioValue.get() == 9:
            txtbox.insert("1.end", rdio_9["text"])

        elif radioValue.get() == 10:
            txtbox.insert("1.end", rdio_10["text"])

            
    #####   GUI(CustomKeywords)   #####
    # ウインドウ作成
    ck_window = tk.Toplevel(root)
    ck_window.geometry("650x430")
    ck_window.resizable(width = False, height = False)

    # ラベル
    ckname_label = tk.Label(ck_window, text = "Custom Keywords:").place(x = 25, y = 10)
    reg_label = tk.Label(ck_window, text = "Registration:").place(x = 25, y = 39)
    tm_label = tk.Label(ck_window, text = "To Main Screen:").place(x = 393, y = 370)

    # キーワード入力欄
    reg_box = tk.Entry(ck_window, width = 60)
    reg_box.place(x = 110, y = 41)
    
    # 保存ボタン
    global save_image
    save_image = tk.PhotoImage(file = "Images/Save.png")
    save_button = tk.Button(ck_window, text =" Save", command = reg_keyword_func,
                           image = save_image, compound = tk.LEFT, width = 60, height = 20)
    save_button.place(x = 547, y = 37)
    # setボタン
    global tomset_image
    tomset_image = tk.PhotoImage(file = "Images/Set.png")
    tomset_button = tk.Button(ck_window, text = " Set", command = ckto_kw_set,
                             image = tomset_image, compound = tk.LEFT,  width = 107, height = 30)
    tomset_button.place(x = 500, y = 360)
    
    # ラベルフレーム(ラジオボタン用)
    ss_frame = ttk.Labelframe(ck_window, text = "Save Slot", paddin = 10)
    ss_frame.place(x = 25, y = 75, relwidth = 0.91)

    # ラジオボタン
    radioValue = tk.IntVar()
    rdio_1 = ttk.Radiobutton(ss_frame, text = config["Custom Keywords1"].get("Keywords1"), variable = radioValue, value = 1) 
    rdio_1.grid(row=0, column=0, sticky = tk.W)
    rdio_2 = ttk.Radiobutton(ss_frame, text = config["Custom Keywords2"].get("Keywords2"), variable = radioValue, value = 2) 
    rdio_2.grid(row=1, column=0, sticky = tk.W)
    rdio_3 = ttk.Radiobutton(ss_frame, text = config["Custom Keywords3"].get("Keywords3"), variable = radioValue, value = 3) 
    rdio_3.grid(row=2, column=0, sticky = tk.W)
    rdio_4 = ttk.Radiobutton(ss_frame, text = config["Custom Keywords4"].get("Keywords4"), variable = radioValue, value = 4) 
    rdio_4.grid(row=3, column=0, sticky = tk.W)
    rdio_5 = ttk.Radiobutton(ss_frame, text = config["Custom Keywords5"].get("Keywords5"), variable = radioValue, value = 5) 
    rdio_5.grid(row=4, column=0, sticky = tk.W)
    rdio_6 = ttk.Radiobutton(ss_frame, text = config["Custom Keywords6"].get("Keywords6"), variable = radioValue, value = 6) 
    rdio_6.grid(row=5, column=0, sticky = tk.W)
    rdio_7 = ttk.Radiobutton(ss_frame, text = config["Custom Keywords7"].get("Keywords7"), variable = radioValue, value = 7) 
    rdio_7.grid(row=6, column=0, sticky = tk.W)
    rdio_8 = ttk.Radiobutton(ss_frame, text = config["Custom Keywords8"].get("Keywords8"), variable = radioValue, value = 8) 
    rdio_8.grid(row=7, column=0, sticky = tk.W)
    rdio_9 = ttk.Radiobutton(ss_frame, text = config["Custom Keywords9"].get("Keywords9"), variable = radioValue, value = 9) 
    rdio_9.grid(row=8, column=0, sticky = tk.W)
    rdio_10 = ttk.Radiobutton(ss_frame, text = config["Custom Keywords10"].get("Keywords10"), variable = radioValue, value = 10) 
    rdio_10.grid(row=9, column=0, sticky = tk.W)

    
# About #
# サブウインドウ定義
def about_Window():
    
    #####   関数(About)  #####
    # リンクを開く
    def callback(url):
        webbrowser.open_new(url)

        
    #####   GUI(About)   #####    
    # ウインドウ作成
    about_window = tk.Toplevel(root)
    about_window.geometry("520x520")
    about_window.resizable(width = False, height = False)

    # 画像指定
    global app_image
    app_image = tk.PhotoImage(file = "Images/abocute_about.png")

    # 画像用キャンパス
    canvas = tk.Canvas(about_window, width = 500, height = 387)
    canvas.place(x = 9, y = 9)
    canvas.create_image(0, 0, image = app_image, anchor = tk.NW)

    # ラベル
    appname_label = tk.Label(about_window, text = "AVOCUTE (Version : 2.0.0)").place(x = 25, y = 405)
    appurl_label = ttk.Label(about_window, text = "http://pc-chem-basics.blog.jp/", foreground = "#3366BB")
    appurl_label.place(x = 25, y = 425)
    appurl_label.bind("<Button-1>", lambda e: callback("http://pc-chem-basics.blog.jp/"))
    about_label = tk.Label(about_window, text = "AVOCUTE is a simple launcher for MOPAC.")
    about_label.place(x = 25, y = 450)

    # 閉じるボタン
    close_button = tk.Button(about_window, text ="Close",  width = 8, command = lambda: about_window.destroy())
    close_button.place(x = 227, y = 480)

# IRC Checker #
##### 関数 ##### 
# FOR*ファイル削除
def for_del():
    file_list = glob.glob("FOR0*")
    for file in file_list:
        os.remove(file)
    
# FOR006(IRC_1SCF)をoutにリネーム,移動
def FOR006_ircrm():
    try:
        # FOR006をoutにリネーム
        if(os.path.exists("FOR006")):
            out_name = os.path.splitext(os.path.basename(ircfile_path))[0] + "_ircf.out"
            os.rename("FOR006", out_name)
            # outファイル移動
            out_move_path = os.path.dirname(ircfile_path)
            out_path = os.path.join(out_move_path, out_name)
            shutil.move(out_name, out_path)
            
            # 全てのタスクの終了を通知            
            tk.messagebox.showinfo("Have fun!!", "Successfully terminated!!")
    except:
        # 同名ファイルが存在する場合エラー通知
        tk.messagebox.showerror("Error", "Error!! [WinError 183]")
              
        
# IRCデータのCSV作成
def csv_create():
    # ファイルの内容をリスト化
    with open(ircfile_path,"r") as outfile:
        data = outfile.readlines()

    # CSVのファイル名
    csv_name = os.path.splitext(os.path.basename(ircfile_path))[0] + "_irc_plot.csv"
    # CSV作成, 見出し追加
    with open(csv_name, "w+") as h:
        writer = csv.writer(h, lineterminator="\n")
        writer.writerow(["POINT"," POTENTIAL", "ENERGY_LOST", "TOTAL", "ERROR", "REF", "", "MOVEMENT"])

    # リストから検索文字列の行番号取得
    for linenum, line in enumerate(data):
        if "OBTAINED" in line:
            finalnum = linenum

            # 検索文字列の3行上を抽出, 単語要素に分割
            words = data[finalnum-3].split()

            # 抽出データをCSVへ書込み
            with open(csv_name,"a+") as datafile:
                writer = csv.writer(datafile, lineterminator="\n")
                writer.writerow(words)
    try:
        # CSV移動
        csv_move_path = os.path.dirname(ircfile_path)
        csv_path = os.path.join(csv_move_path, csv_name)
        shutil.move(csv_name, csv_path)
    except:
        # 同名ファイルが存在する場合エラー通知
        tk.messagebox.showerror("Error", "Error!! [WinError 183]")
        
# IRCの最終座標をoutputで出力(1SCF計算)
def mopac_1scf_run():
    # MOPAC(exe)が存在するかチェック
    if os.path.exists("mopac606.exe"):
            # MOPAC実行
            subprocess.run(["mopac606.exe", "FOR005"])
            
            if os.path.exists("FOR006"):
                # FOR006をoutにリネーム,移動
                FOR006_ircrm()
                # IRCデータCSV作成
                csv_create()
                # FOR*ファイル削除
                for_del()
            else:
                # FOR*ファイル削除
                for_del()
                # 座標がない場合はエラー表示
                tk.messagebox.showerror("Error", "No coordinates!!")
                
    # MOPACが見つからない場合はエラー表示
    else:
        tk.messagebox.showerror("Error", "No mopac.exe!!")

# IRCデータからFOR005作成(最終座標取得)
def scfdat_create():
    # ファイルの内容をリスト化
    with open(ircfile_path,"r") as outfile:
        data = outfile.readlines()
        
    # リストから検索文字列の行番号取得
    for linenum, line in enumerate(data):
        if "OBTAINED" in line:
            finalnum = linenum

            # 検索文字列の4行目から最終行までリスト内要素を抽出
            words = data[(finalnum+4):]

            # FOR005作成
            with open("FOR005","w+") as inputfile:
                inputfile.write("1SCF\nTitle\n\n")

            # 抽出データの書込み
            with open("FOR005","a+") as datafile:
                # ファイルへ書込み
                for line in words:
                    # 最初の空行で処理を終了
                    if line == "\n":
                        datafile.write("\n")
                        break
                    else:
                        datafile.write(line)
    # 1SCF計算               
    mopac_1scf_run()

# IRCデータのファイルパス取得
def irc_getpath(): 
    global ircfile_path
    typ = [("MOPAC","*.out")]
    ircfile_path = filedialog.askopenfilename(filetypes = typ)
    
    if len(ircfile_path) != 0:
        # IRCデータからFOR005作成(最終座標取得)
        scfdat_create()
    else:
        tk.messagebox.showwarning("Warning", "Output file not selected!!")

    
###########################
#####   Main Screen  ##### 

#####   関数(Main)   #####
# datファイル保存
def save_Input(data):
    global dat_path
    # ファイル選択ダイアログの表示
    typ = [('MOPAC','*.dat')] 
    dir = 'C:\\pg'
    dat_path = tk.filedialog.asksaveasfilename(filetypes = typ, initialdir = dir, defaultextension = "dat")
    
    if len(dat_path) != 0:
        # ファイルが選択された場合データを書き込み
        f = open(dat_path, "w")
        f.write(data)
        f.close()
        # FOR005へdatの内容コピー
        FOR005 = open("FOR005","w+")
        FOR005.write(data)
        FOR005.close()
    else:
        # ステータス更新
        statusbar["text"] = "  No input file!!"

# datファイル読込み
def load_Input():
    global loaddat_path
    # ファイル選択ダイアログの表示
    typ = [('MOPAC','*.dat')] 
    dir = 'C:\\pg'
    loaddat_path = tk.filedialog.askopenfilename(filetypes = typ, initialdir = dir, defaultextension = "dat")
    
    if len(loaddat_path) != 0:
        # ファイルが選択された場合データを書き込み
        
        # 既存のtxtboxデータ消去
        txtbox.delete("1.0","end")
        # ファイルのリスト取得
        with open(loaddat_path) as f:
            txtbox.insert("1.end", f.read())
            
        # ステータス更新
        statusbar["text"] = "  Let's Start!!"        
        
    else:
        # ステータス更新
        statusbar["text"] = "  No input file!!"

# arc一時ファイルをテキストボックスへ挿入
def read_arc():
    # 既存のtxtboxデータ消去
    txtbox.delete("1.0","end")
    # ファイルのリスト取得
    with open("temp.txt") as f:
        txtbox.insert("1.end", f.read())

    # ステータス更新
    statusbar["text"] = "  Let's Start!!"  

# arc一時ファイル作成、読込み
def create_arc_temp():
    # 一時fileへの書込み先指定
    temp_file = open("temp.txt","w+")
    # ファイルのリスト取得
    with open(loadarc_path,"r") as outfile_arc:
        arc_list = outfile_arc.readlines()
    
    try:
        # リストから検索文字列の行番号取得
        for linenum_arc, line_arc in enumerate(arc_list):
            if "FINAL" in line_arc:
                finalnum_arc = linenum_arc

        # 検索文字列の1行目から最終行までリスト内要素を抽出
        coord__arc = arc_list[(finalnum_arc +1):]
        # ファイルへ書込み
        for arc_line in coord__arc:
            temp_file.write(arc_line)
        # ファイルを閉じる
        temp_file.close()
        # arc一時ファイルをテキストボックスへ挿入
        read_arc()
    except:
        # ステータス更新
        statusbar["text"] = "  no coordinates!!" 
    
# arcファイル読込み
def load_arc():
    global loadarc_path
    # ファイル選択ダイアログの表示
    typ = [('MOPAC','*.arc')] 
    dir = 'C:\\pg'
    loadarc_path = tk.filedialog.askopenfilename(filetypes = typ, initialdir = dir, defaultextension = "arc")
    
    if len(loadarc_path) != 0:
        # ファイルが選択された場合データを書き込み
        
        # arc一時ファイル作成、読込み
        create_arc_temp()
    else:
        # ステータス更新
        statusbar["text"] = "  No input file!!"        
        
# AUXチェク
def aux_chk(data):
    # 互換モードがon, AUXを含む場合は削除
    if mode_chkvalue.get() and "AUX" in data:
        data = data.replace("AUX", "")
        save_Input(data)
        
    else:
        save_Input(data)

# テキストボックス(dat作成欄)からテキストを取得
def write_button_func():
    # ステータス更新
    statusbar["text"] = "  Let's Start!!"
    # テキストボックスから文字列取得
    data = txtbox.get("1.0", "end")
    
    # AUXチェク
    aux_chk(data)

# outputエラー確認
def stts_chk():
    # outputのパス取得
    output_path = os.path.splitext(dat_path)[0] + ".out"
    try:
        # outputの内容をリスト化
        with open(output_path,"r") as outfile_stts:
            stts_list = outfile_stts.readlines()
        # outputからエラーが無いか確認
        if " == MOPAC DONE ==\n" in stts_list:
            # ステータス更新
            statusbar["text"] = "  MOPAC terminated!! == MOPAC DONE =="
        else:
            # ステータス更新
            statusbar["text"] = "  MOPAC terminated!!  Error: Check the output!!" 
    except:
        # ステータス更新
        statusbar["text"] = "  MOPAC terminated!!  Error: Check the input!!"  

# FOR023をmgfにリネーム,移動
def FOR023_rm():
    # FOR023をmgfにリネーム
    if(os.path.exists("FOR023")):
        mgf_name = os.path.splitext(os.path.basename(dat_path))[0] + ".mgf"
        os.rename("FOR023", mgf_name)
        # arcファイル移動
        mgf_move_path = os.path.dirname(dat_path)
        mgf_path = os.path.join(mgf_move_path, mgf_name)
        shutil.move(mgf_name, mgf_path)
        
# FOR012をarcにリネーム,移動
def FOR012_rm():
    # FOR012をarcにリネーム
    if(os.path.exists("FOR012")):
        arc_name = os.path.splitext(os.path.basename(dat_path))[0] + ".arc"
        os.rename("FOR012", arc_name)
        # arcファイル移動
        arc_move_path = os.path.dirname(dat_path)
        arc_path = os.path.join(arc_move_path, arc_name)
        shutil.move(arc_name, arc_path)
        
# FOR006をoutにリネーム,移動
def FOR006_rm():
    # FOR006をoutにリネーム
    if(os.path.exists("FOR006")):
        out_name = os.path.splitext(os.path.basename(dat_path))[0] + ".out"
        os.rename("FOR006", out_name)
        # outファイル移動
        out_move_path = os.path.dirname(dat_path)
        out_path = os.path.join(out_move_path, out_name)
        shutil.move(out_name, out_path)
        
# MOPACでdatファイルを実行
def mopac_run():
    # MOPAC(exe)が存在するかチェック
    if os.path.exists("mopac606.exe"):
        # ステータスから文字列取得
        status = statusbar.cget("text")

        if status != "  No input file!!":
            # ステータス更新
            statusbar["text"] = "  wait..." 
            root.update_idletasks()
            # MOPAC実行
            subprocess.run(["mopac606.exe", "FOR005"])
            
            # ホームディレクトリで上書き保存するとエラーが発生するので回避
            try:
                # 終了後にFOR006をoutにリネーム,移動
                FOR006_rm()
                # 終了後にFOR012をarcにリネーム,移動
                FOR012_rm()
                # 終了後にFOR023をmgfにリネーム,移動
                FOR023_rm()

                # 終了後に一時ファイル削除
                file_list = glob.glob("FOR0*")
                for file in file_list:
                    os.remove(file)

                # 終了後にステータス更新
                stts_chk()
            except:
                # FOR*ファイル削除
                file_list = glob.glob("FOR0*")
                for file in file_list:
                    os.remove(file)
                # ステータス更新
                statusbar["text"] = "  Error!! [WinError 183]"                
    
    # MOPACが見つからないときはエラー表示
    else:
        messagebox.showerror('Error', 'No mopac.exe!!')

# キーワード設定値をテキストボックスへ印字
def kw_set():
    # ステータス更新
    statusbar["text"] = "  Let's Start!!"
    # 既存の値削除
    txtbox.delete("1.0","1.end")
    
    # 値の挿入
    for combo in (calc_dict[calc_combo.get()], with_dict[with_combo.get()],
              chg_dict[chg_combo.get()], multi_dict[multi_combo.get()]):
        txtbox.insert("1.end", combo)
    for cb in (v1, v2, v3, v4, v5):
        txtbox.insert("1.end", cb.get())

# Custom Keywords起動時のini確認
def config_chk():
    # 設定ファイルが存在するかチェック
    if os.path.exists("settings.ini"):
        
        # 存在する場合はCustom Keywordsを開く
        create_CK_Window()
        
    # 存在しない場合はエラー表示
    else:
        messagebox.showerror("Error", "No settings.ini!!")

# メニューバーから一時ファイル削除
def clear_yesno():
    # ダイアログで削除のyes/no
    ret = messagebox.askyesno("Clear Temporary Files", 'Are you sure you want to delete the FOR*** Files?')
    # yesなら削除
    if ret == True:
        file_list = glob.glob("FOR0*")
        for file in file_list:
            os.remove(file)

# JSmol(web)へアクセス
def jsmol_web():
    jsmol_url = "https://chemapps.stolaf.edu/jmol/jmol.php?source=?"
    webbrowser.open(jsmol_url)
    
# HP(配布ページ)へアクセス
def hp_web():
    hp_url = "http://pc-chem-basics.blog.jp/archives/23981733.html"
    webbrowser.open(hp_url)

# Help(MOPAC)を開く
def mopac_help():
    mopac_help1 = "Help/MOPAC6_Manual.txt"
    subprocess.run(["start", mopac_help1], shell=True)

# Help(MOPACキーワード)を開く
def mopackey_help():
    mopac_help2 = "Help/MOPAC6_Keywords.txt"
    subprocess.run(["start", mopac_help2], shell=True)
    

#####   GUI(Main)   #####
# ウインドウ作成
root = tk.Tk()
root.geometry("650x650")
root.title("Avocute Ver.2.0.0")
root.iconphoto(True, tk.PhotoImage(file = "Images/taskbar_icon.png"))
root.resizable(width = False, height = False)

# メニューバー作成
menubar = tk.Menu(root)
root.configure(menu = menubar)
# File_Menu
filemenu = tk.Menu(menubar, tearoff=0)
menubar.add_cascade(label = "File", menu = filemenu)
# File_Menuの内容
open_menu = tk.Menu(menubar, tearoff = 0)
filemenu.add_cascade(label = "Open...", menu = open_menu)
open_menu.add_command(label = "Input File (*.dat)", command = load_Input)
open_menu.add_command(label = "Arc File (*.arc)", command = load_arc)
filemenu.add_command(label = "Save as... (*.dat)", command = write_button_func)
filemenu.add_command(label = "Clear Temporary Files", command = lambda: clear_yesno())
filemenu.add_separator()
filemenu.add_command(label = "Exit", command = lambda: root.destroy())
# Tools_Menu
toolsmenu = tk.Menu(menubar, tearoff=0)
menubar.add_cascade(label = "Tools", menu = toolsmenu)
# Tools_Menuの内容
toolsmenu.add_command(label = "Batch Launcher", command = create_BL_Window)
toolsmenu.add_command(label = "Auto TS Searcher", command = create_ATS_Window)
toolsmenu.add_command(label = "IRC Checker", command = irc_getpath)
toolsmenu.add_separator()
toolsmenu.add_command(label = "JSmol", command = jsmol_web)
# Help_Menu
helpmenu = tk.Menu(menubar, tearoff=0)
menubar.add_cascade(label = "Help", menu = helpmenu)
# Help_Menuの内容
helpmenu.add_command(label = "About", command = about_Window)
helpmenu.add_separator()
helpmenu.add_command(label = "MOPAC6 Manual", command = mopac_help)
helpmenu.add_command(label = "MOPAC6 Keywords", command = mopackey_help)
helpmenu.add_separator()
helpmenu.add_command(label = "Website", command = hp_web)

# ラベル作成
mopac_input_label = tk.Label(root, text = "MOPAC Input:").place(x = 20, y = 190)
calc_label = tk.Label(root, text = "Calculation:").place(x = 20, y = 20)
with_label = tk.Label(root, text = "With:").place(x = 300, y = 20)
chg_label = tk.Label(root, text = "Charge:").place(x = 20, y = 50)
multi_label = tk.Label(root, text = "Multiplicity:").place(x = 265, y = 50)

# ラベルフレーム作成(オプション用)
options_frame = ttk.Labelframe(root, text = "Options", paddin = 10)
options_frame.place(x = 20, y = 100)

# コンボボックス作成
calc_dict = {"Equilibrium Geometry": "EF", "Frequencies": "FORCE",
             "Transition State": "TS", "Saddle Point": "SADDLE", "IRC": "IRC=1 LARGE=50"}
calc_combo = ttk.Combobox(root, state = "readonly", values = list(calc_dict.keys()))
calc_combo.place(x = 100, y = 20)
calc_combo.current(0)

with_dict = {"AM1": " AM1", "PM3": " PM3"}
with_combo = ttk.Combobox(root, state ="readonly", width = 8, values = list(with_dict.keys()))
with_combo.place(x = 345, y = 20)
with_combo.current(0)

chg_dict = {"Dication": " CHARGE=+2", "Cation": " CHARGE=+1", "Neutral": "", "Anion": " CHARGE=-1", "Dianion": " CHARGE=-2"}
chg_combo = ttk.Combobox(root, state = "readonly", width = 8, values = list(chg_dict.keys()))
chg_combo.place(x = 100, y = 50)
chg_combo.current(2)

multi_dict = {"Singlet": "", "Doublet": " DOUBLET", "Triplet": " TRIPLET"}
multi_combo = ttk.Combobox(root, state = "readonly", width = 8, values = list(multi_dict.keys()))
multi_combo.place(x = 345, y = 50)
multi_combo.current(0)

# チェックボタン作成(オプション用)
v1 = tk.StringVar()
cb1 = ttk.Checkbutton(options_frame, text = "PRECISE", variable = v1, onvalue = " PRECISE", offvalue = "", padding = 10)
cb1.grid(row=0, column=0)

v2 = tk.StringVar()
cb2 = ttk.Checkbutton(options_frame, text = "GEO-OK", variable = v2, onvalue = " GEO-OK", offvalue = "", padding = 10)
cb2.grid(row=0, column=1)

v3 = tk.StringVar()
cb3 = ttk.Checkbutton(options_frame, text = "XYZ", variable = v3, onvalue = " XYZ", offvalue = "", padding = 10)
cb3.grid(row=0, column=2)

v4 = tk.StringVar()
cb4 = ttk.Checkbutton(options_frame, text = "UHF", variable = v4, onvalue = " UHF", offvalue = "", padding = 10)
cb4.grid(row=0, column=3)

v5 = tk.StringVar()
cb5 = ttk.Checkbutton(options_frame, text = "GRAPH", variable = v5, onvalue = " GRAPH", offvalue = "", padding = 10)
cb5.grid(row=0, column=4)

# 互換性チェクボタン
mode_chkvalue = tk.BooleanVar()
mode_chkvalue.set(True)
mode_chk = ttk.Checkbutton(root, text = "Avogadro Mode", variable = mode_chkvalue)
mode_chk.place(x = 20, y = 548)

# Custom Keywordsボタン作成
ps_image = tk.PhotoImage(file = "Images/Preset.png")
ck_button = tk.Button(root, text = " Preset", command = config_chk,  image = ps_image, compound = tk.LEFT,  width = 107, height = 30)
ck_button.place(x = 490, y = 20)
# Setボタン作成
set_image = tk.PhotoImage(file = "Images/Set.png")
set_button = tk.Button(root, text = " Set", command = kw_set, image = set_image, compound = tk.LEFT,  width = 107, height = 30)
set_button.place(x = 490, y = 133)
# Runボタン作成(datを保存して実行)
run_image = tk.PhotoImage(file = "Images/Run.png")
run_button = tk.Button(root, text = " Run", image = run_image, compound = tk.LEFT,  width = 540, height = 30,
                     command = lambda:[write_button_func(), mopac_run()])
run_button.place(x = 38, y = 575)

# dat編集用テキストボックス作成
txtbox = tk.Text(root, width = 83, height = 25)
txtbox.place(x = 20, y = 215)

# スクロールバー作成
yscroll = tk.Scrollbar(root, orient = tk.VERTICAL, command = txtbox.yview)
yscroll.place(x = 605, y = 215, relheight = 0.51)
txtbox["yscrollcommand"] = yscroll.set

# ステータスバー作成
statusbar = tk.Label(root, bd = 1, text =  "  Let's Start!!", relief = tk.SUNKEN, anchor = tk.W)
statusbar.pack(side = tk.BOTTOM, fill=tk.X)


# ウインドウ状態の維持
root.mainloop()