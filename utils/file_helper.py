def handle_uploaded_file(complete_path, file):
    try:
        with open(complete_path, 'wb+') as destination:
            for chunk in file.chunks():
                destination.write(chunk)
        return True
    except Exception as e:
        return False