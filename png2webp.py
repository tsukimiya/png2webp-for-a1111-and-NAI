import glob
import os
import piexif
import piexif.helper
from PIL import Image, PngImagePlugin
from datetime import datetime
from pywintypes import Time
import argparse
from tqdm import tqdm
import concurrent.futures
import threading
import time

# Windowsの場合
on_windows = os.name == 'nt'
if on_windows:
    import win32file
    import win32con

# WEBP品質
WEBP_QUALITY = 100
# 画像形式
IMG_INPUT_FORMAT = 'PNG'
IMG_OUTPUT_FORMAT = 'WEBP'
# 画像拡張子
IMG_INPUT_FILENAME_EXT = 'png'
IMG_OUTPUT_FILENAME_EXT = 'webp'

def format_size(size):
    # Helper function to format size in KB, MB, GB
    for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
        if size < 1024:
            return f"{size:.2f} {unit}"
        size /= 1024

def convert_image(file, delete_original, lossless):
    file_name = os.path.splitext(os.path.basename(file))[0]
    output_file_name = file_name + '.' + IMG_OUTPUT_FILENAME_EXT
    output_file_path = os.path.join(os.path.dirname(file), output_file_name)
    output_file_abspath = os.path.abspath(output_file_path)

    def get_png_info(file):
        try:
            # 画像を開く
            img = Image.open(file)

            # PngImagePluginの情報を取得
            png_info = img.info

            # ファイルを閉じる
            img.close()

            return png_info

        except Exception as e:
            print(f"画像を開けませんでした。: {e}")
            return None

    # PNGファイルからpnginfoを取得
    png_info = get_png_info(file)

    # 画像を開く
    image = Image.open(file)

    # 日時情報を取得
    access_time   = os.path.getatime(file) # アクセス日時
    modify_time   = os.path.getmtime(file) # 更新日時

    if on_windows:
        creation_time = os.path.getctime(file) # 作成日時

    # 元のファイルサイズを取得
    original_size = os.path.getsize(file)

    # WEBPに変換
    if lossless:
        image.save(output_file_path, IMG_OUTPUT_FORMAT, lossless=True, quality=WEBP_QUALITY)
    else:
        image.save(output_file_path, IMG_OUTPUT_FORMAT, quality=WEBP_QUALITY)

    # 画像を閉じる
    image.close()

    # 新しいファイルサイズを取得
    new_size = os.path.getsize(output_file_path)
    size_reduction = original_size - new_size

    # WEBPファイルにExifデータ（PNG Info）を保存する
    if png_info is not None:
        # pnginfoの各項目を改行区切りで連結
        png_info_data = ""
        for key, value in png_info.items():
            if key == 'parameters':
                # Automatic1111形式の場合
                png_info_data += f"{value}\n"
            else:
                # NovelAI形式の場合
                png_info_data += f"{key}: {value}\n"

        png_info_data = png_info_data.rstrip()

        # Exifデータを作成
        exif_dict = {"Exif": {piexif.ExifIFD.UserComment: piexif.helper.UserComment.dump(png_info_data or '', encoding='unicode')}}

        # Exifデータをバイトに変換
        exif_bytes = piexif.dump(exif_dict)

        # Exifデータを挿入して新しい画像を保存
        piexif.insert(exif_bytes, output_file_path)

    else:
        print("PNG情報を取得できませんでした。")

    # 日付情報の設定
    # WEBPファイルのハンドルを取得（Windowsのみ）
    if on_windows:
        handle = win32file.CreateFile(
            output_file_path,
            win32con.GENERIC_WRITE,
            win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE | win32con.FILE_SHARE_DELETE,
            None,
            win32con.OPEN_EXISTING,
            0,
            None
        )

        # WEBPファイルに元画像の作成日時、アクセス日時、更新日時を設定
        win32file.SetFileTime(handle, Time(creation_time), Time(access_time), Time(modify_time))

        # ハンドルを閉じる
        handle.Close()

    # 他のプラットフォームではアクセス日時と更新日時を設定
    os.utime(output_file_path, (access_time, modify_time))

    # 元のPNGファイルを削除
    if delete_original:
        os.remove(file)

    return size_reduction

def convert_images_in_directory(directory, delete_original, lossless):
    # 画像を配列に格納
    files = glob.glob(os.path.join(directory, '**', '*.' + IMG_INPUT_FILENAME_EXT), recursive=True)
    total_files = len(files)
    total_size_reduction = 0

    # Limit the number of concurrent threads
    max_threads = min(4, os.cpu_count() - 1)  # Adjust the number of threads as needed
    semaphore = threading.Semaphore(max_threads)

    def thread_task(file, pbar):
        nonlocal total_size_reduction
        with semaphore:
            result = convert_image(file, delete_original, lossless)
            time.sleep(0.1)  # Add a small sleep interval to reduce CPU usage
            total_size_reduction += result
            pbar.set_description(f"Processing images (Total size reduction: {format_size(total_size_reduction)})")
            pbar.update(1)
            return result

    with tqdm(total=total_files, desc="Processing images") as pbar:
        with concurrent.futures.ThreadPoolExecutor(max_workers=max_threads) as executor:
            futures = [executor.submit(thread_task, file, pbar) for file in files]
            for future in concurrent.futures.as_completed(futures):
                try:
                    future.result()
                except Exception as e:
                    print(f"Error processing file: {e}")

    print(f"Processed {total_files} files.")
    print(f"Total size reduction: {format_size(total_size_reduction)}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Convert PNG images to WEBP format in a specified directory.')
    parser.add_argument('directory', type=str, help='The directory containing PNG images to convert.')
    parser.add_argument('--delete', action='store_true', help='Delete original PNG files after conversion.')
    parser.add_argument('--lossless', action='store_true', help='Convert images to lossless WEBP format.')
    args = parser.parse_args()
    convert_images_in_directory(args.directory, args.delete, args.lossless)