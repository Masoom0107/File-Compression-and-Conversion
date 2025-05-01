import subprocess

# List of libraries to install
libraries = [
    #'Flask',
    #'docx',
    #'Pillow',
    #'python-pptx',
    #'openpyxl',
    'aspose-python',
    #'moviepy',
]

# Install each library using pip
for library in libraries:
    try:
        subprocess.check_call(['pip', 'install', library])
        print(f'Successfully installed {library}')
    except subprocess.CalledProcessError:
        print(f'Failed to install {library}')
