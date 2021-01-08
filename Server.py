import socket
import FileProcess

ONE_CONNECTION_ONLY = (True)
fileInput = "server/File Receive.xlsx"
fileOutput = "server/File Giám thị.xls"

port = 1218
sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
host = "192.168.0.101"
sock.bind((host, port))
sock.listen()
print("File Server started...")
print(f"[LISTENING] Server is listening on {host}")

while True:
    conn, addr = sock.accept()
    print(f"Accepted connection from {addr}")
    with open(fileInput, "wb") as file:
        print("File open")
        print("Receiving data...")
        i = 0
        while True:
            data = conn.recv(1024)
            print(i)
            i+=1
            if not data:
                break
            file.write(data)

    print("File receive")

    conn.close()
    
    conn, addr = sock.accept()
    FileProcess.readFromExcel(fileInput)
    FileProcess.supervisorInRoom()
    FileProcess.supervisorOutRoom()
    FileProcess.writeToExcel(fileOutput)
    FileProcess.connectToDatabase()

    with open(fileOutput, "rb") as file:
        data = file.read(1024)
        
        i = 0
        while data:
            conn.send(data)
            #print(f"Sent {data!r}")
            print(i)
            i+=1
            data = file.read(1024)
    print("File send")

    conn.close()




#while True:
#    conn, addr = sock.accept()
#    print(f"Accepted connection from {addr}")
#    data = conn.recv(1024)
#    print(f"Server received {data}")
#    with open(filename, "rb") as file:
#        data = file.read(1024)
#        while data:
#            conn.send(data)
#            print(f"Sent {data!r}")
#            data = file.read(1024)

#    print("File sent complete.")
#    conn.close()
#    if (ONE_CONNECTION_ONLY):
#        break
#sock.shutdown(1)
#sock.close()