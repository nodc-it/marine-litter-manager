# marine-litter-manager

![logo](https://user-images.githubusercontent.com/8235122/159733958-270b8390-8f27-4b66-b543-c9a8bb1c8e31.png)

Marine Litter Manager (MLM) is a Python data formatting tool that can be used to generate:

* EMODnet beach litter format
* EMODnet seafloor trawlings litter format

This is done following the specifications of the official guidelines published by [EMODnet Chemistry](https://www.emodnet-chemistry.eu/). It is available for Linux and Windows.

The software is available for download at the following link: https://www.emodnet-chemistry.eu/marinelitter/manager

The user manual is available at the following link: https://doi.org/10.13120/21addf37-7e82-4a55-b040-3d3d87115ac0

The specific video tutorials are available at the following link: https://www.emodnet-chemistry.eu/help/littervideotutorial

## Available functions
* Seafloor Litter formatting
* Beach Litter formatting
* The surveys plots (for both litter standards)
* The params plots (for both litter standards)
* Dictionary search engine
* Links

## How to build an executable
### Windows
```batch
pip install pyinstaller
pip install -r requirements.txt

pyinstaller --noconfirm --onefile --windowed --icon "./resources/logo.ico" --add-data "./resources;." --hidden-import "PIL._tkinter_finder"  mlm.py
```
If you are using Anaconda you must exclude some modules with the `--exclude-module` option: `--exclude-module scikit-learn,PyQt5,PyQt4,2to3,IPython,Jinja2,pycparser,scipy`
### Linux
```bash
pip install pyinstaller
pip install -r requirements.txt

pyinstaller --noconfirm --onefile --windowed --icon ./resources/logo.ico --add-data ./resources:. --hidden-import PIL._tkinter_finder  mlm.py
```

## Screenshots

![main](https://user-images.githubusercontent.com/8235122/159734629-f884e917-d92a-42a5-a955-17ad0255bc97.png)

![10](https://user-images.githubusercontent.com/8235122/159734650-9ac89b69-9382-445a-bcf4-8116d1b63a61.png)

![11](https://user-images.githubusercontent.com/8235122/159734666-6fe0d3a4-ddcc-4fc4-b8e2-d459687c5c42.png)

![12](https://user-images.githubusercontent.com/8235122/159734688-8b3c9aeb-8991-4e6b-b8cf-487617dc5337.png)

![13](https://user-images.githubusercontent.com/8235122/159734701-459d13cd-7e43-4608-9bd8-8850d2db6a5b.png)

![14](https://user-images.githubusercontent.com/8235122/159734722-6df2b8ed-59a7-4749-8e83-56fe7ee44fe4.png)

![15](https://user-images.githubusercontent.com/8235122/159734734-8dd2b9ca-c079-4aff-bbdb-2a20b9ee7f64.png)

![16](https://user-images.githubusercontent.com/8235122/159734742-a71b70ae-4c12-4639-aac0-eaf2d53c523e.png)

![17](https://user-images.githubusercontent.com/8235122/159734753-f0dde669-ae63-4a51-99e8-465b6ee4d84c.png)

![18](https://user-images.githubusercontent.com/8235122/159734768-67761dac-04fe-49b4-a689-b55ddd226478.png)

![19](https://user-images.githubusercontent.com/8235122/159734784-6b86a121-aac6-4b51-9a58-a64c1da5f1e3.png)

![20](https://user-images.githubusercontent.com/8235122/159734799-fb8a0740-ad11-4ca5-8d15-838ffbbb97f0.png)

![21](https://user-images.githubusercontent.com/8235122/159734822-6d62988a-733f-462f-b307-6b3644ae9a78.png)

![22](https://user-images.githubusercontent.com/8235122/159734848-b378df4c-8f3c-47ee-ae84-62b11038d35d.png)
