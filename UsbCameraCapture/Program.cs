using CommandLine;
using NetMQ;
using NetMQ.Sockets;
using System;
using System.Runtime.Versioning;
using System.Text.Json;
using System.Collections.Generic;

namespace UsbCameraCapture
{
    class Program
    {
        public class Options
        {
            [Option('p', "port", Required = true, HelpText = "Set ZeroMQ port number.")]
            public int Port { get; set;  }
        }

        public class ZeroMQMessage
        {
            public string MessageId { get; set; }

            public string JsonString { get; set; }
        }

        public class ZeroMQDevice
        {
            public string Name { get; set; }

            public string DevicePath { get; set; }
        }

        public class ZeroMQVideoInfo
        {
            public string DevicePath { get; set; }

            public int Width { get; set; }

            public int Height { get; set; }

            public short Bitrate { get; set; }

            public long AvgTimePerFrame { get; set; }
        }

        public class ZeroMQResult
        {
            public bool Result { get; set; }
        }

        public class ZeroMQFrame
        {
            public bool Result { get; set; }

            public string Timestamp { get; set; }

            public string DataType { get; set; }

            public int[] Shape { get; set; }
        }

        [SupportedOSPlatform("windows")]
        static void Main(string[] args)
        {
            int port = 0;
            Parser.Default.ParseArguments<Options>(args).WithParsed<Options>(o =>
            {
                port = o.Port;
            });

            var capture = new DirectShowCapture();

            using (var responseSocket = new ResponseSocket())
            {
                // コマンドライン引数からポート番号を取得する
                responseSocket.Bind($"tcp://*:{port}");

                var isCancellation = false;
                while (!isCancellation)
                {
                    var message = JsonSerializer.Deserialize<ZeroMQMessage>(responseSocket.ReceiveFrameString());

                    switch(message.MessageId)
                    {
                        case "start_capture":
                            {
                                // --
                                // DevicePath, Width, Height, Bitrate, FPSが必要
                                // --

                                var videoInfo = JsonSerializer.Deserialize<ZeroMQVideoInfo>(message.JsonString);
                                var result = capture.Start(videoInfo.DevicePath, videoInfo.Width, videoInfo.Height, videoInfo.Bitrate, videoInfo.AvgTimePerFrame);

                                var resultMessage = new ZeroMQResult() { Result = result };
                                responseSocket.SendFrame(JsonSerializer.Serialize(resultMessage));
                            }
                            break;

                        case "stop_capture":
                            {
                                // -- 
                                // 引数無し
                                // --

                                capture.Stop();

                                responseSocket.SendFrameEmpty();
                            }
                            break;

                        case "get_capture_devices":
                            {
                                // --
                                // 引数無し
                                // --

                                var devices = DirectShowCapture.GetCaptureDevices();
                                var deviceMessages = new List<ZeroMQDevice>();
                                foreach (var device in devices)
                                {
                                    deviceMessages.Add(new ZeroMQDevice() { Name = device.Name, DevicePath = device.DevicePath });
                                }

                                responseSocket.SendFrame(JsonSerializer.Serialize(deviceMessages));
                            }
                            break;

                        case "get_video_infos":
                            {
                                // --
                                // DevicePath, DeviceNameが必要
                                // --

                                var device = JsonSerializer.Deserialize<ZeroMQDevice>(message.JsonString);
                                var videoInfos = DirectShowCapture.GetVideoInfos(device.DevicePath, device.Name);

                                responseSocket.SendFrame(JsonSerializer.Serialize(videoInfos));
                            }
                            break;

                        case "get_frame":
                            {
                                // --
                                // 引数無し
                                // --

                                string timestamp;
                                byte[] image;
                                int? height, width;

                                var result = capture.GetFrame(out timestamp, out image, out height, out width);
                                if (result)
                                {
                                    var frameInfo = new ZeroMQFrame() { Result = true, Timestamp = timestamp, DataType = "uint8", Shape = new int[] { height.Value, width.Value, 4 } };
                                    responseSocket.SendMoreFrame(JsonSerializer.Serialize(frameInfo)).SendFrame(image);
                                }
                                else
                                {
                                    var frameInfo = new ZeroMQFrame() { Result = false };
                                    responseSocket.SendFrame(JsonSerializer.Serialize(frameInfo));
                                }
                            }
                            break;

                        case "thumbnail":
                            {
                                // --
                                // 引数無し
                                // --

                                string timestamp;
                                byte[] image;
                                int? height, width;

                                var result = capture.GetThumbnail(out timestamp, out image, out height, out width);
                                if (result)
                                {
                                    var frameInfo = new ZeroMQFrame() { Result = true, Timestamp = timestamp, DataType = "uint8", Shape = new int[] { height.Value, width.Value, 4 } };
                                    responseSocket.SendMoreFrame(JsonSerializer.Serialize(frameInfo)).SendFrame(image);
                                }
                                else
                                {
                                    var frameInfo = new ZeroMQFrame() { Result = false };
                                    responseSocket.SendFrame(JsonSerializer.Serialize(frameInfo));
                                }
                            }
                            break;

                        case "exit":
                            {
                                // --
                                // 引数無し
                                // --

                                // プログラムを終了します
                                responseSocket.SendFrameEmpty();

                                isCancellation = true;
                            }
                            break;

                        default:
                            {
                                // 上記以外のメッセージにはPING-PONGの返信として「PONG」を送る
                                responseSocket.SendFrame("PONG");
                            }
                            break;
                    }
                }
            }
        }
    }
}
