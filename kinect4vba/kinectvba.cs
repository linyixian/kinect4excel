using System;
using Microsoft.Kinect;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;

namespace kinect4vba
{
    [ClassInterface(ClassInterfaceType.AutoDual)]

    public class kinectvba
    {

        private KinectSensor kinect;

        //コンストラクタ

        public kinectvba()
        {
            if (KinectSensor.KinectSensors.Count == 0)
                throw new Exception("kinectが接続されていません。");

            kinect = KinectSensor.KinectSensors[0];
        }

        //kinectの使用開始
        public void Start()
        {
            StartKinect(kinect);
        }

        public void Stop()
        {
            StopKinect(kinect);
        }

        private void StartKinect(KinectSensor kinect)
        {
            //RGBカメラを有効
            kinect.ColorStream.Enable();

            kinect.Start();
        
        }

        private void StopKinect(KinectSensor kinect)
        {
            if (kinect != null)
            {
                if (kinect.IsRunning)
                {
                    //kinectの停止
                    kinect.Stop();
                    kinect.Dispose();
                }
            }
        }

        //Bitmapをファイルに保存
        public Boolean GetPict(string path, ref string message)
        {
            if (kinect.IsRunning==false)
            {
                StartKinect(kinect);
            }

            using (ColorImageFrame colorFrame=kinect.ColorStream.OpenNextFrame(100))
            {
                if (colorFrame == null)
                {
                    return false;
                }

                try
                {
                    byte[] colorPixel = new byte[colorFrame.PixelDataLength];
                    colorFrame.CopyPixelDataTo(colorPixel);

                    //ピクセルデータをビットマップに変換
                    Bitmap bitmap =
                    new Bitmap(kinect.ColorStream.FrameWidth,
                               kinect.ColorStream.FrameHeight,
                               PixelFormat.Format32bppRgb);

                    Rectangle rect = new Rectangle(0, 0, bitmap.Width, bitmap.Height);

                    BitmapData data = bitmap.LockBits(rect, ImageLockMode.WriteOnly,
                        PixelFormat.Format32bppRgb);

                    Marshal.Copy(colorPixel, 0, data.Scan0, colorPixel.Length);

                    bitmap.UnlockBits(data);

                    //ファイルに保存
                    bitmap.Save(path, ImageFormat.Jpeg);

                    return true;
                }
                catch (Exception ex)
                {
                    message += ex.Message;
                    return false;
                }
            }

        }

        //Bitmapをクリップボードに保存
        public Boolean GetPict2Clip()
        {
            if (kinect.IsRunning==false)
            {
                StartKinect(kinect);
            }

            using (ColorImageFrame colorFrame=kinect.ColorStream.OpenNextFrame(100))
            {
                if (colorFrame == null)
                {
                    return false;
                }
                    
                byte[] colorPixel = new byte[colorFrame.PixelDataLength];
                colorFrame.CopyPixelDataTo(colorPixel);

                //ピクセルデータをビットマップに変換
                Bitmap bitmap =
                new Bitmap(kinect.ColorStream.FrameWidth,
                           kinect.ColorStream.FrameHeight,
                           PixelFormat.Format32bppRgb);

                Rectangle rect = new Rectangle(0, 0, bitmap.Width, bitmap.Height);

                BitmapData data = bitmap.LockBits(rect, ImageLockMode.WriteOnly,
                    PixelFormat.Format32bppRgb);

                Marshal.Copy(colorPixel, 0, data.Scan0, colorPixel.Length);

                bitmap.UnlockBits(data);

                //クリップボードに保存
                DataObject dataobj = new DataObject();
                dataobj.SetImage(bitmap);
                Clipboard.SetDataObject(dataobj, true);

                return true;
            }
        }


        //kinectの動作状況
        public Boolean IsRunning(KinectSensor kinect)
        {
            return kinect.IsRunning;
        }


        //チルト初期位置
        public void TiltInit()
        {
            if (kinect.IsRunning == false)
            {
               return;
            }

            kinect.ElevationAngle = 0;

            //連続で動かすと例外が発生するので1秒間Sleep
            System.Threading.Thread.Sleep(1000);
        }

        //チルト上向き
        public void TiltUp()
        {
            if (kinect.IsRunning == false)
            {
                return;
            }

            if (kinect.ElevationAngle < kinect.MaxElevationAngle - 2)
            {
                kinect.ElevationAngle  += 2;
            }

            //連続で動かすと例外が発生するので1秒間Sleep
            System.Threading.Thread.Sleep(1000);
        }

        //チルト下向き
        public void TiltDown()
        {

            if (kinect.IsRunning == false)
            {
               return;
            }

            if (kinect.ElevationAngle > kinect.MinElevationAngle + 2)
            {
                kinect.ElevationAngle -= 2;
            }

            //連続で動かすと例外が発生するので1秒間Sleep
            System.Threading.Thread.Sleep(1000);
        }
    }
}
