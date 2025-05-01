import subprocess

# List of libraries to install
libraries_to_install = [
    #"Flask==2.3.3",
    #"Flask-SocketIO==5.3.5",
    #"python-docx==0.8.11",
    #"moviepy==1.0.3",
    #"Pillow==9.5.0",
    #"python-pptx==0.6.22",
    #"openpyxl==3.1.2",
    "aspose-pdf==23.7.0"
]

# Install each library using pip
for library in libraries_to_install:
    try:
        subprocess.check_call(["pip", "install", library])
        print(f"Successfully installed {library}")
    except subprocess.CalledProcessError:
        print(f"Error installing {library}")

# Libraries that are part of Python's standard library (no need to install)
standard_libraries = [
    "threading",
    "logging",
    "zipfile",
    "tempfile",
    "os",
    "io",
    "shutil",
    "sys",
    "time"
]

print("Libraries that are part of Python's standard library (no need to install):")
for library in standard_libraries:
    print(library)
