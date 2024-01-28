import PyInstaller.__main__

PyInstaller.__main__.run([
    '--name=%s' % "marksBot",
    '--onefile',
    '--windowed',
    # '--add-binary=%s' % './path/to/your/assets:assets',
    # '--icon=%s' % './path/to/your/icon.ico',
    'src',
])
