using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using Microsoft.Kinect;
using Microsoft.Speech.AudioFormat;
using Microsoft.Speech.Recognition;
using System.IO;

//鼠标操作
using System.Runtime.InteropServices;
using Coding4Fun.Kinect;
using Coding4Fun.Kinect.Wpf;
/*
using Microsoft.Win32;
using OFFICECORE = Microsoft.Office.Core;
using POWERPOINT = Microsoft.Office.Interop.PowerPoint;
using System.Collections;
//*/


namespace KinectCtrlPPT
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {


        private KinectSensor _kinect;
        Choices voicecommands = new Choices();

        bool isWindowsClosing = false; //窗口是否正在关闭中
        const int MaxSkeletonTrackingCount = 6; //最多同时可以跟踪的用户数
        Skeleton[] allSkeletons = new Skeleton[MaxSkeletonTrackingCount];//骨骼跟踪



        private void InitKinect()
        {
            if (KinectSensor.KinectSensors.Count > 0)
            {
                //选择第一个Kinect设备
                _kinect = KinectSensor.KinectSensors[0];

                if (_kinect == null)
                {
                    return;
                }

                //启用深度摄像头和彩色摄像头，为了获得更好的映射效果，彩色摄像头的分辨率恰好是深度摄像头分辨率的2倍
                _kinect.DepthStream.Enable(DepthImageFormat.Resolution640x480Fps30);     //获取深度信息流
                _kinect.ColorStream.Enable(ColorImageFormat.RgbResolution1280x960Fps12);  //获取彩色图像流

                var parameters = new TransformSmoothParameters
                {
                    Smoothing = 0.5f,
                    Correction = 0.5f,
                    Prediction = 0.5f,
                    JitterRadius = 0.05f,
                    MaxDeviationRadius = 0.04f
                };
                _kinect.SkeletonStream.Enable(parameters);                               //获取骨骼跟踪数据

                _kinect.SkeletonFrameReady += new EventHandler<SkeletonFrameReadyEventArgs>(kinect_SkeletonFrameReady);//预定事件

                try
                {
                    //显示彩色图像摄像头
                    kinectColorViewer1.Kinect = _kinect;

                    //启动
                    _kinect.Start();

                    //语音命令播放PPT
                    PPTPlayViaVoice();
                }
                catch (System.IO.IOException)
                {
                    MessageBox.Show("Kinect device is not found!");
                }
            }
            else
            {
                MessageBox.Show("Kinect device is not found!");
            }
        }

        void kinect_SkeletonFrameReady(object sender, SkeletonFrameReadyEventArgs e)
        {
            //骨骼跟踪状态提示
            labelIsSkeletonTracked.Visibility = System.Windows.Visibility.Hidden;

            if (isWindowsClosing)
            {
                return;
            }

            //Get a skeleton
            Skeleton s = getClosetSkeleton(e);

            if (s == null)
            {
                return;
            }

            if (s.TrackingState != SkeletonTrackingState.Tracked)
            {
                return;
            }

            //提示用户骨骼跟踪就绪,可以进行操作
            if (s.TrackingState == SkeletonTrackingState.Tracked)
                labelIsSkeletonTracked.Visibility = System.Windows.Visibility.Visible;

            //映射键盘事件,演示PPT
            presentPowerPoint(s);
        }



        // private const double JumpDiffThreadhold = 0.05; //跳跃的落差阀值，单位米
        // private double headPreviousPosition = 2.0; //初始值，一般人不会有2米身高 跳跃动作留空，未映射


        private bool isNextGestureActive = false;
        private bool isLastGestureActive = false;
        private bool isBlackScreenActive = false;

        private bool isForwardGestureActive = false;

        private const double ArmStretchedThreshold = 0.4; //演示PPT时，手臂水平伸展的阈值，单位米
        private const double ArmRaisedThreshold = 0.2;    //演示PPT时，手臂垂直举起的阈值，单位米

        private const double HandFowardThreshold = -0.4;   //演示PPT时，手掌前推的阈值，单位米

        private void presentPowerPoint(Skeleton s)
        {
            SkeletonPoint head = s.Joints[JointType.Head].Position;
            SkeletonPoint leftshoulder = s.Joints[JointType.ShoulderLeft].Position;
            SkeletonPoint rightshoulder = s.Joints[JointType.ShoulderRight].Position;

            SkeletonPoint rightWrist = s.Joints[JointType.WristRight].Position;

            SkeletonPoint leftHand = s.Joints[JointType.HandLeft].Position;
            SkeletonPoint rightHand = s.Joints[JointType.HandRight].Position;

            bool isRightHandRaised = (rightHand.Y - rightshoulder.Y) > ArmRaisedThreshold;
            bool isLeftHandRaised = (leftHand.Y - leftshoulder.Y) > ArmRaisedThreshold;

            bool isRightHandStretched = (rightHand.X - rightshoulder.X) > ArmStretchedThreshold;
            bool isLeftHandStretched = (leftshoulder.X - leftHand.X) > ArmStretchedThreshold;


            /*//        测试手掌前推阈值
                        float test_z = rightWrist.Z - rightshoulder.Z;
                        if (test_z < -0.4)
                        {
                            MessageBox.Show(test_z.ToString());
                            return;            
                        }            
            //*/
            bool isRightHandForward = (rightWrist.Z - rightshoulder.Z) < HandFowardThreshold;


            //使用状态变量，避免多次重复发送键盘事件
            //右手水平伸展开
            if (isRightHandStretched)
            {
                if (!isLastGestureActive && !isNextGestureActive)
                {
                    KeyboardToolkit.Keyboard.Type(Key.Right);
                    isNextGestureActive = true;

                    //System.Windows.Forms.SendKeys.SendWait("{Right}");
                }
            }
            else
            {
                isNextGestureActive = false;
            }
            //左手水平伸展开
            if (isLeftHandStretched)
            {
                if (!isLastGestureActive && !isNextGestureActive)
                {
                    KeyboardToolkit.Keyboard.Type(Key.Left);
                    isLastGestureActive = true;

                    //System.Windows.Forms.SendKeys.SendWait("{Left}");
                }
            }
            else
            {
                isLastGestureActive = false;
            }

            ////双手同时举起，控制PPT时则让屏幕变黑
            if (isLeftHandRaised && isRightHandRaised)
            {
                if (!isBlackScreenActive)
                {
                    KeyboardToolkit.Keyboard.Type(Key.B);
                    isBlackScreenActive = true;

                    //System.Windows.Forms.SendKeys.SendWait("{B}");
                }
            }
            else
            {
                isBlackScreenActive = false;
            }

            //进行标注

            // 右手前推表示单击按下，后撤表示单击释放
            if (isRightHandForward)
            {
                if (!isForwardGestureActive)
                {
                    mouse_event((int)(MouseEventFlags.LeftDown | MouseEventFlags.Absolute), 0, 0, 0, IntPtr.Zero);
                    //MessageBox.Show("press down the mouse!");
                    isForwardGestureActive = true;
                    
                }                
            }
            else
            {
                mouse_event((int)(MouseEventFlags.LeftUp | MouseEventFlags.Absolute), 0, 0, 0, IntPtr.Zero);
                //MessageBox.Show("press up the mouse!");
                isForwardGestureActive = false;
            }

            // 移动鼠标位置
            /*         var hand = (leftHand.Y > rightHand.Y)
                                       ? s.Joints[JointType.HandLeft]
                                       : s.Joints[JointType.HandRight];
           //*/
            var hand = s.Joints[JointType.HandLeft];
            if (hand.TrackingState != JointTrackingState.Tracked)
                return;

            //获取当前屏幕的宽高
            int ScreenWidth = (int)SystemParameters.PrimaryScreenWidth;
            int ScreenHeight = (int)SystemParameters.PrimaryScreenHeight;

            float posX = hand.ScaleTo(ScreenWidth, ScreenHeight, 0.2f, 0.2f).Position.X;
            float posY = hand.ScaleTo(ScreenWidth, ScreenHeight, 0.2f, 0.2f).Position.Y;
            //bool is_mark = false;
            SetCursorPos((int)posX, (int)posY);

        }


        Skeleton getClosetSkeleton(SkeletonFrameReadyEventArgs e)
        {
            using (SkeletonFrame skeletonFrameData = e.OpenSkeletonFrame())
            {
                if (skeletonFrameData == null)
                {
                    return null;
                }

                skeletonFrameData.CopySkeletonDataTo(allSkeletons);

                //Linq语法，查找离Kinect最近的、被跟踪的骨骼
                Skeleton closestSkeleton = (from s in allSkeletons
                                            where s.TrackingState == SkeletonTrackingState.Tracked &&
                                                  s.Joints[JointType.Head].TrackingState == JointTrackingState.Tracked
                                            select s).OrderBy(s => s.Joints[JointType.Head].Position.Z)
                                    .FirstOrDefault();

                return closestSkeleton;
            }
        }

        private void stopKinect(KinectSensor sensor)
        {
            if (sensor != null)
            {
                if (sensor.IsRunning)
                {
                    //stop sensor 
                    sensor.Stop();

                    //关闭音频流，如果当前已打开的话
                    if (sensor.AudioSource != null)
                    {
                        sensor.AudioSource.Stop();
                    }
                }
            }
        }


        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            labelIsSkeletonTracked.Visibility = System.Windows.Visibility.Hidden;

            try
            {
                //monitoring for Kinect's status
                // 监听Kinect的状态
                KinectSensor.KinectSensors.StatusChanged += new EventHandler<StatusChangedEventArgs>(KinectSensors_StatusChanged);

                //loop through all the Kinects attached to this PC, and start the first that is connected without an error.
                // 寻找所有连接在电脑上的Kinect设备，并启动第一个连接正常的设备
                foreach (KinectSensor kinect in KinectSensor.KinectSensors)
                {
                    if (kinect.Status == KinectStatus.Connected)
                    {
                        this._kinect = kinect;
                        break;
                    }
                }

                if (KinectSensor.KinectSensors.Count == 0)
                    MessageBox.Show("No Kinect found");
                else
                    this.InitKinect();

                //设置应用程序手型光标
                Mouse.OverrideCursor = Cursors.Hand;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        // Kinect 状态改变触发事件
        private void KinectSensors_StatusChanged(object sender, StatusChangedEventArgs e)
        {
            switch (e.Status)
            {
                case KinectStatus.Connected:
                    if (_kinect == null)
                    {
                        _kinect = e.Sensor;
                        this.InitKinect();
                    }
                    break;
                case KinectStatus.Disconnected:
                    if (_kinect == e.Sensor)
                    {
                        releaseResources();
                        MessageBox.Show("Kinect disconnected");
                    }
                    break;
                case KinectStatus.NotReady:
                    break;
                case KinectStatus.NotPowered:
                    if (_kinect == e.Sensor)
                    {
                        releaseResources();
                        MessageBox.Show("Kinect powered off");
                    }
                    break;
                default:
                    MessageBox.Show("Unhandled Status: " + e.Status);
                    break;
            }
        }

        private void releaseResources()
        {

            if (this.voicecommands != null)
            {
                this.speechrecoengine.SpeechRecognized -= speechrecoengine_SpeechRecognized;
                this._kinect.AudioSource.Stop();
                this.voicecommands = null;
            }

            if (_kinect != null)
            {
                _kinect.SkeletonFrameReady -= this.kinect_SkeletonFrameReady;
                _kinect.Stop();
                _kinect = null;
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            isWindowsClosing = true;
            stopKinect(_kinect);
        }

        //=================================================================================
        //语音播放PPT
        private SpeechRecognitionEngine speechrecoengine;

        private static RecognizerInfo GetKinectRecognizer()
        {
            Func<RecognizerInfo, bool> matchingFunc = r =>
            {
                string value;
                r.AdditionalInfo.TryGetValue("Kinect", out value);
                return "True".Equals(value, StringComparison.InvariantCultureIgnoreCase) && "en-US".Equals(r.Culture.Name, StringComparison.InvariantCultureIgnoreCase);
            };
            return SpeechRecognitionEngine.InstalledRecognizers().Where(matchingFunc).FirstOrDefault();
        }


        private void PPTPlayViaVoice()
        {
            // 等待4秒钟时间，让Kinect传感器初始化启动完成
            System.Threading.Thread.Sleep(1000);

            // 获取Kinect音频对象
            KinectAudioSource audiosource = _kinect.AudioSource;
            audiosource.EchoCancellationMode = EchoCancellationMode.None; // 本示例中关闭“回声抑制模式”
            audiosource.AutomaticGainControlEnabled = false; // 启用语音命令识别需要关闭“自动增益”

            RecognizerInfo recoinfo = GetKinectRecognizer();

            if (recoinfo == null)
            {
                MessageBox.Show("Could not find Kinect speech recognizer.");
                return;
            }

            speechrecoengine = new SpeechRecognitionEngine(recoinfo.Id);

            // 添加语音命令 ok-开始播放 / thanks-停止播放
            
            voicecommands.Add("okay");
            voicecommands.Add("thank you");

            //荧光标注前的准备工作
            voicecommands.Add("mark");
            voicecommands.Add("stop mark");
            voicecommands.Add("discard");
            voicecommands.Add("keep");

            var grambuilder = new GrammarBuilder { Culture = recoinfo.Culture };

            // 创建语法对象                                
            grambuilder.Append(voicecommands);

            //根据语言区域，创建语法识别对象
            var grammar = new Grammar(grambuilder);

            // 将这些语法规则加载进语音识别引擎    
            speechrecoengine.LoadGrammar(grammar);

            // 注册事件：有效语音命令识别、疑似识别、无效识别
            speechrecoengine.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(speechrecoengine_SpeechRecognized);
            speechrecoengine.SpeechHypothesized += new EventHandler<SpeechHypothesizedEventArgs>(speechrecoengine_SpeechHypothesized);
            speechrecoengine.SpeechRecognitionRejected += new EventHandler<SpeechRecognitionRejectedEventArgs>(speechrecoengine_SpeechRecognitionRejected);

            // 初始化并启动 Kinect音频流
            Stream s = audiosource.Start();
            speechrecoengine.SetInputToAudioStream(
                s, new SpeechAudioFormatInfo(EncodingFormat.Pcm, 16000, 16, 1, 32000, 2, null));

            // 异步开启语音识别引擎，可识别多次
            speechrecoengine.RecognizeAsync(RecognizeMode.Multiple);
        }

        void speechrecoengine_SpeechRecognitionRejected(object sender, SpeechRecognitionRejectedEventArgs e)
        {
            //throw new NotImplementedException();
            //MessageBox.Show("SpeechRecognitionRejected");
        }

        void speechrecoengine_SpeechHypothesized(object sender, SpeechHypothesizedEventArgs e)
        {
            //throw new NotImplementedException();
            //MessageBox.Show("SpeechHypothesized");
        }

        void speechrecoengine_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            //throw new NotImplementedException();
            //语音识别信心度超过70%
            if (e.Result.Confidence >= 0.7)
            {
                string voicecommand = e.Result.Text.ToLower();
                if (voicecommand == "okay")
                {
                    KeyboardToolkit.Keyboard.Type(Key.F5);
                    labelstart.Visibility = System.Windows.Visibility.Hidden;
                }
                else if (voicecommand == "thank you")
                {
                    KeyboardToolkit.Keyboard.Type(Key.Escape);
                    labelesc.Visibility = System.Windows.Visibility.Hidden;
                }
                else if (voicecommand == "mark")
                {
                    //鼠标右击
                    try
                    {
                        RightClick();
                        KeyboardToolkit.Keyboard.Type(Key.O);
                        KeyboardToolkit.Keyboard.Type(Key.B);
                    }
                    catch
                    {
                        MessageBox.Show("标注准备失败！");
                    }
                    //end
                    // */
                    //KeyboardToolkit.Keyboard.Type(Key.O);
                }
                else if (voicecommand == "stop mark")
                {
                    try
                    {
                        RightClick();
                        KeyboardToolkit.Keyboard.Type(Key.O);
                        KeyboardToolkit.Keyboard.Type(Key.A);
                    }
                    catch
                    {
                        MessageBox.Show("停止标注失败！");
                    }
                }
                else if (voicecommand == "discard")
                {
                    KeyboardToolkit.Keyboard.Type(Key.D);
                }
                else if (voicecommand == "keep")
                {
                    KeyboardToolkit.Keyboard.Type(Key.K);
                }
            }
        }

        // 鼠标操作  也可引用Kinect.Toolbox.Cursor
        #region
        [DllImport("User32")]
        public extern static void mouse_event(int dwFlags, int dx, int dy, int dwData, IntPtr dwExtraInfo);

        [DllImport("User32")]
        public extern static void SetCursorPos(int x, int y);

        [DllImport("User32")]
        public extern static bool GetCursorPos(out POINT pt);

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;
        }

        public enum MouseEventFlags
        {
            Move = 0x0001,
            LeftDown = 0x0002,
            LeftUp = 0x0004,
            RightDown = 0x0008,
            RightUp = 0x0010,
            MiddleDown = 0x0020,
            MiddleUp = 0x0040,
            Wheel = 0x0800,
            Absolute = 0x8000
        }
        #endregion
        POINT CursorPosition = new POINT();

        private void RightClick()
        {
            POINT p = new POINT();

            GetCursorPos(out p);
            CursorPosition = p;
            SetCursorPos(CursorPosition.X, CursorPosition.Y);

            try
            {
                mouse_event((int)(MouseEventFlags.RightDown | MouseEventFlags.Absolute), 0, 0, 0, IntPtr.Zero);
                mouse_event((int)(MouseEventFlags.RightUp | MouseEventFlags.Absolute), 0, 0, 0, IntPtr.Zero);
            }
            catch
            {
                MessageBox.Show("右击动作发生异常");
            }
            finally
            {
                SetCursorPos(p.X, p.Y);
            }
        }

        //====================打开PPT文件===========================


        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            /*
            //创建一个打开文件式的对话框  
            OpenFileDialog ofd = new OpenFileDialog();
            //设置这个对话框的起始打开路径  
            ofd.InitialDirectory = @"C:\Users";
            //设置打开的文件的类型，注意过滤器的语法  
            ofd.Filter = "PPT文件（.ppt）|*.ppt";
            String filePath = ofd.FileName;
            //调用ShowDialog()方法显示该对话框，该方法的返回值代表用户是否点击了确定按钮  
            if (ofd.ShowDialog() == true)
            {
                PPTOpen(filePath);

            }
            
            else
            {
                MessageBox.Show("没有选择PPT");
            }
            //*/
        }

        private void MenuItem_Click_1(object sender, RoutedEventArgs e)
        {
            //PPTClose();
        }
/*
        #region=========基本的参数信息=======
        POWERPOINT.Application objApp = null;
        POWERPOINT.Presentation objPresSet = null;
        POWERPOINT.SlideShowWindows objSSWs;
        POWERPOINT.SlideShowTransition objSST;
        POWERPOINT.SlideShowSettings objSSS;
        POWERPOINT.SlideRange objSldRng;
        bool bAssistantOn;
        double pixperPoint = 0;
        double offsetx = 0;
        double offsety = 0;
        #endregion
        #region===========操作方法==============
        /// <summary>
        /// 打开PPT文档并播放显示。
        /// </summary>
        /// <param name="filePath">PPT文件路径</param>
        public void PPTOpen(string filePath)
        {
            //防止连续打开多个PPT程序.
            if (this.objApp != null) { return; }
            try
            {
                objApp = new POWERPOINT.Application();
                //以非只读方式打开,方便操作结束后保存.
                objPresSet = objApp.Presentations.Open(filePath, OFFICECORE.MsoTriState.msoFalse, OFFICECORE.MsoTriState.msoFalse, OFFICECORE.MsoTriState.msoFalse);
                //Prevent Office Assistant from displaying alert messages:
                bAssistantOn = objApp.Assistant.On;
                objApp.Assistant.On = false;
                objSSS = this.objPresSet.SlideShowSettings;
                objSSS.Run();
            }
            catch
            {
                this.objApp.Quit();
            }
        }
        /// <summary>
        /// 自动播放PPT文档.
        /// </summary>
        /// <param name="filePath">PPTy文件路径.</param>
        /// <param name="playTime">翻页的时间间隔.【以秒为单位】</param>
        public void PPTAuto(string filePath, int playTime)
        {
            //防止连续打开多个PPT程序.
            if (this.objApp != null) { return; }
            objApp = new POWERPOINT.Application();
            objPresSet = objApp.Presentations.Open(filePath, OFFICECORE.MsoTriState.msoCTrue, OFFICECORE.MsoTriState.msoFalse, OFFICECORE.MsoTriState.msoFalse);
            // 自动播放的代码（开始）
            int Slides = objPresSet.Slides.Count;
            int[] SlideIdx = new int[Slides];
            for (int i = 0; i < Slides; i++) { SlideIdx[i] = i + 1; };
            objSldRng = objPresSet.Slides.Range(SlideIdx);
            objSST = objSldRng.SlideShowTransition;
            //设置翻页的时间.
            objSST.AdvanceOnTime = OFFICECORE.MsoTriState.msoCTrue;
            objSST.AdvanceTime = playTime;
            //翻页时的特效!
            objSST.EntryEffect = POWERPOINT.PpEntryEffect.ppEffectCircleOut;
            //Prevent Office Assistant from displaying alert messages:
            bAssistantOn = objApp.Assistant.On;
            objApp.Assistant.On = false;
            //Run the Slide show from slides 1 thru 3.
            objSSS = objPresSet.SlideShowSettings;
            objSSS.StartingSlide = 1;
            objSSS.EndingSlide = Slides;
            objSSS.Run();
            //Wait for the slide show to end.
            objSSWs = objApp.SlideShowWindows;
            while (objSSWs.Count >= 1) System.Threading.Thread.Sleep(playTime * 100);
            this.objPresSet.Close();
            this.objApp.Quit();
        }
        /// <summary>
        /// PPT下一页。
        /// </summary>
        public void NextSlide()
        {
            if (this.objApp != null)
                this.objPresSet.SlideShowWindow.View.Next();
        }
        /// <summary>
        /// PPT上一页。
        /// </summary>
        public void PreviousSlide()
        {
            if (this.objApp != null)
                this.objPresSet.SlideShowWindow.View.Previous();
        }
        
        /// <summary>
        /// 关闭PPT文档。
        /// </summary>
        public void PPTClose()
        {
            //装备PPT程序。
            if (this.objPresSet != null)
            {
                //判断是否退出程序,可以不使用。
                //objSSWs = objApp.SlideShowWindows;
                //if (objSSWs.Count >= 1)
                //{
                    if (MessageBox.Show("是否保存修改的笔迹!", "提示", MessageBoxButton.OKCancel) == MessageBoxResult.OK)
                        this.objPresSet.Save();
                //}
                //this.objPresSet.Close();
            }
            if (this.objApp != null)
                this.objApp.Quit();
            GC.Collect();
        }
        #endregion

 //*/       

    }
}