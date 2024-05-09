using System;
using System.IO.Ports;
using System.Runtime.InteropServices;
using System.Threading;

class Program
{
    const int waitTime = 5 * 1000;

    [DllImport("user32.dll")]
    public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
    [DllImport("user32.dll")]
    static extern bool SetCursorPos(int X, int Y);
    [DllImport("user32.dll")]
    public static extern bool GetCursorPos(out POINT lpPoint);
    [StructLayout(LayoutKind.Sequential)]
    public struct POINT
    {
        public int X;
        public int Y;
    }

    static SerialPort serialPort;

    static void Main(string[] args)
    {
        string SerialPath = args[0];
        int BaudRate = int.Parse(args[1]);
        // string SerialPath = "COM3";
        // int BaudRate = 9600;

        Console.WriteLine($"Serial Path: {SerialPath}");
        Console.WriteLine($"Baud Rate: {BaudRate}");

        Connect();

        void Connect()
        {
            Console.WriteLine();
            Console.WriteLine($"Connecting in {waitTime / 1000} seconds");
            Thread.Sleep(waitTime);

            serialPort = new SerialPort(SerialPath, BaudRate);
            serialPort.DataReceived += SerialPort_DataReceived;

            try
            {
                Console.WriteLine("Connecting to serial port");
                serialPort.Open();
                serialPort.DtrEnable = true;
                Console.WriteLine("Connected to serial port");
                while (serialPort.IsOpen) { };
                Console.WriteLine("Connection to serial port closed");
                Connect();
            }
            catch (Exception err)
            {
                Console.WriteLine($"Serial Error: {err.Message}");
                Connect();
            }
        }
    }

    static void SerialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
    {
        try
        {
            string data = serialPort.ReadLine();
            Console.WriteLine($"Received data: {data}");

            string[] dataSplit = data.Split(':');
            string type = dataSplit[0];
            string value = dataSplit[1].Replace("\r", "");

            if (type == "k")
            {
                // Key press
                SendKeys(value);
            }
            else
            if (type == "kd")
            {
                // Key down
                SendKeysDown(value);
            }
            else
            if (type == "ku")
            {
                // Key up
                SendKeysUp(value);
            }
            else
            if (type == "m")
            {
                // Set Mouse
                string[] movePositions = value.Split(',');
                int moveX = int.Parse(movePositions[0]);
                int moveY = int.Parse(movePositions[1]);
                SetCursorPos(moveX, moveY);
                //MoveMouse(moveX, moveY);
            }
            if (type == "mm")
            {
                // Move Mouse
                string[] movePositions = value.Split(',');
                int moveX = int.Parse(movePositions[0]);
                int moveY = int.Parse(movePositions[1]);
                MoveMouse(moveX, moveY);
            }
        }
        catch (Exception err)
        {
            Console.WriteLine($"Error reading line: {err.Message}");
        }
    }

    static void SendKeys(string keys)
    {
        foreach (char key in keys)
        {
            byte keyCode = (byte)char.ToUpper(key);
            keybd_event(keyCode, 0, 0x0000, UIntPtr.Zero); // Press down
            keybd_event(keyCode, 0, 0x0002, UIntPtr.Zero); // Release
        }
    }

    static void SendKeysDown(string keys)
    {
        foreach (char key in keys)
        {
            byte keyCode = (byte)char.ToUpper(key);
            keybd_event(keyCode, 0, 0x0000, UIntPtr.Zero); // Press down
        }
    }

    static void SendKeysUp(string keys)
    {
        foreach (char key in keys)
        {
            byte keyCode = (byte)char.ToUpper(key);
            keybd_event(keyCode, 0, 0x0002, UIntPtr.Zero); // Release
        }
    }

    static void MoveMouse(int X, int Y)
    {
        POINT currentPosition;
        GetCursorPos(out currentPosition);

        SetCursorPos(currentPosition.X + X, currentPosition.Y - Y);
    }
}
