import socket
import tkinter
from tkinter import filedialog, Text
import os

fileInput = ""
root = tkinter.Tk()


def addApp():
    for widget in frame.winfo_children():
        widget.destroy()

    fileInput = filedialog.askopenfilename(initialdir="/", title="Select File",
                                            filetypes=(("excel", "*.xlsx"), ("all files", "*.*")))

    print(fileInput)

    fileOutPut = "client/File Giám Thị.xls"
    sock = socket.socket()
    host = socket.gethostname()
    port = 1218
    sock.connect((host, port))

    with open(fileInput, "rb") as file:
        data = file.read(1024)
        while data:
            sock.send(data)
            print(f"Sent {data!r}")
            data = file.read(1024)

    print("File sent complete.")

    sock.close()
    sock = socket.socket()
    host = socket.gethostname()
    port = 1218
    sock.connect((host, port))

    with open(fileOutPut, "wb") as file:
        while True:
            data = sock.recv(1024)
            print(f"data={data}")
            if not data:
                break
            file.write(data)

    print("Got the file")
    sock.close()
    print("Connection is closed")

canvas = tkinter.Canvas(root, height=700, width=700, bg="#263D42")
canvas.pack()

frame = tkinter.Frame(root, bg="white")
frame.place(relwidth=0.8, relheight=0.8, relx=0.1, rely=0.1)

openFile = tkinter.Button(root, text="Open File", padx=10,
                            pady=5, fg="white", bg="#263D42", command=addApp)
openFile.pack()

#runApps = tkinter.Button(root, text="Run Apps", padx=10,
#                            pady=5, fg="white", bg="#263D42", command=runApps)
#runApps.pack()

root.mainloop()

