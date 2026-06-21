from utils.uploads import save_excel_upload


def save_excel_upload_to_session(file_storage, session_data, session_key, upload_folder, *, prefix):
    file_path = save_excel_upload(file_storage, upload_folder, prefix=prefix)
    session_data[session_key] = file_path
    return file_path
