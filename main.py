import sys
import os
import shutil
import platform
import subprocess

def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    input_dir = os.path.join(base_dir, "keikakusho")
    link_source_dir = os.path.join(base_dir, "link")
    output_base_dir = os.path.join(base_dir, "output")

    # 1. 元データ(Word)の確認
    word_files = [f for f in os.listdir(input_dir) if f.endswith('.docx') and not f.startswith('~')]
    if not word_files:
        print("エラー: keikakushoフォルダにWordファイルが見つかりません。")
        return
    
    original_word_name = word_files[0]
    word_name_no_ext = os.path.splitext(original_word_name)[0]
    
    # 2. 出力フォルダの設定（元データの名前にする）
    save_dir = os.path.join(output_base_dir, word_name_no_ext)
    if os.path.exists(save_dir):
        shutil.rmtree(save_dir)
    os.makedirs(save_dir)

    # 3. Wordとリンク資料をコピー
    target_word_path = os.path.join(save_dir, original_word_name)
    shutil.copy2(os.path.join(input_dir, original_word_name), target_word_path)
    
    # linkフォルダの中身を丸ごとコピー（階層維持）
    link_items = os.listdir(link_source_dir)
    link_data = []
    for item in link_items:
        src_path = os.path.join(link_source_dir, item)
        dst_path = os.path.join(save_dir, item)
        if os.path.isdir(src_path):
            shutil.copytree(src_path, dst_path)
            # フォルダ内の全ファイルをリンク対象としてスキャン
            for root, _, files in os.walk(dst_path):
                for file in files:
                    if file.startswith('.'): continue
                    file_no_ext = os.path.splitext(file)[0]
                    rel_path = os.path.relpath(os.path.join(root, file), save_dir).replace("\\", "/")
                    link_data.append((file_no_ext, rel_path))
        else:
            shutil.copy2(src_path, dst_path)
            file_no_ext = os.path.splitext(item)[0]
            link_data.append((file_no_ext, item))

    # 4. OS別のWord操作
    os_name = platform.system()
    if os_name == "Darwin": # Mac
        applescript = f'''
        tell application "Microsoft Word"
            open POSIX file "{os.path.abspath(target_word_path)}"
            set doc to active document
        '''
        for text, path in link_data:
            applescript += f'''
            set findRange to text object of doc
            repeat
                execute find (find object of findRange) find text "{text}"
                if found of (find object of findRange) is true then
                    make new hyperlink object at doc with properties {{text object:findRange, address:"{path}"}}
                    collapse range findRange direction collapse end
                else
                    exit repeat
                end if
            end repeat
            '''
        applescript += 'save doc\nclose doc\nend tell'
        subprocess.run(["osascript", "-e", applescript])
    
    elif os_name == "Windows":
        import win32com.client
        word_app = win32com.client.Dispatch("Word.Application")
        doc = word_app.Documents.Open(os.path.abspath(target_word_path))
        for text, path in link_data:
            word_app.Selection.HomeKey(6)
            find = word_app.Selection.Find
            while find.Execute(text):
                doc.Hyperlinks.Add(Anchor=word_app.Selection.Range, Address=path)
        doc.Save()
        doc.Close()
        word_app.Quit()

    print(f"完了！ output/{word_name_no_ext} フォルダを確認してください。")

if __name__ == "__main__":
    main()
