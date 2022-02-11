import socket
import tkinter
from tkinter import filedialog

fileInput = ""
root = tkinter.Tk()
root.title("Supervisor Management")


def addApp():
    fileInput = filedialog.askopenfilename(initialdir="/", title="Select File",
                                           filetypes=(("excel", "*.xlsx"), ("all files", "*.*")))

    print(fileInput)

    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    host = "192.168.16.1"
    port = 1218
    sock.connect((host, port))

    with open(fileInput, "rb") as file:
        data = file.read(1024)
        while data:
            sock.send(data)
            # print(f"Sent {data!r}")
            data = file.read(1024)

    print("File sent complete.")
    sock.close()

    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    host = "192.168.16.1"
    port = 1218
    sock.connect((host, port))
    file = filedialog.asksaveasfile(mode='wb', defaultextension=".xls")
    while True:
        data = sock.recv(1024)
        # print(f"data={data}")
        if not data:
            break
        file.write(data)
    file.close()

    print("Got the file")
    sock.close()
    print("Connection is closed")


canvas = tkinter.Canvas(root, height=240, width=300, bg="#263D42")
canvas.pack()

openFile = tkinter.Button(root, text="Open File", padx=10,
                          pady=5, fg="white", bg="#263D42", command=addApp)
openFile.pack()

root.mainloop()
