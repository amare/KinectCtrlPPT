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



namespace KinectCtrlPPT
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        
       
        private KinectSensor kinect;

        bool isWindowsClosing = false; //窗口是否正在关闭中
        const int MaxSkeletonTrackingCount = 6; //最多同时可以跟踪的用户数
        Skeleton[] allSkeletons = new Skeleton[MaxSkeletonTrackingCount];//骨骼跟踪



        private void startKinect()
        {
            if (KinectSensor.KinectSensors.Count > 0)
            {
                //选择第一个Kinect设备
                kinect = KinectSensor.KinectSensors[0];

                if (kinect == null)
                {
                    return;
                }

                //启用深度摄像头和彩色摄像头，为了获得更好的映射效果，彩色摄像头的分辨率恰好是深度摄像头分辨率的2倍
                kinect.DepthStream.Enable(DepthImageFormat.Resolution640x480Fps30);     //获取深度信息流
                kinect.ColorStream.Enable(ColorImageFormat.RgbResolution1280x960Fps12);  //获取彩色图像流

                var parameters = new TransformSmoothParameters
                {
                    Smoothing = 0.5f,
                    Correction = 0.5f,
                    Prediction = 0.5f,
                    JitterRadius = 0.05f,
                    MaxDeviationRadius = 0.04f
                };
                kinect.SkeletonStream.Enable(parameters);                               //获取骨骼跟踪数据

                kinect.SkeletonFrameReady += new EventHandler<SkeletonFrameReadyEventArgs>(kinect_SkeletonFrameReady);//预定事件

                try
                {
                    //显示彩色图像摄像头
                    kinectColorViewer1.Kinect = kinect;

                    //启动
                    kinect.Start();

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
        private const double ArmStretchedThreadhold4PPT = 0.4; //演示PPT时，手臂水平伸展的阀值，单位米
        private const double ArmRaisedThreshhold = 0.2;        //演示PPT时，手臂垂直举起的阀值，单位米

        private void presentPowerPoint(Skeleton s)
        {
            SkeletonPoint head = s.Joints[JointType.Head].Position;
            SkeletonPoint leftshoulder = s.Joints[JointType.ShoulderLeft].Position;
            SkeletonPoint rightshoulder = s.Joints[JointType.ShoulderRight].Position;

            SkeletonPoint leftHand = s.Joints[JointType.HandLeft].Position;
            SkeletonPoint rightHand = s.Joints[JointType.HandRight].Position;

            bool isRightHandRaised = (rightHand.Y - rightshoulder.Y) > ArmRaisedThreshhold;
            bool isLeftHandRaised = (leftHand.Y - leftshoulder.Y) > ArmRaisedThreshhold;

            bool isRightHandStretched = (rightHand.X - rightshoulder.X) > ArmStretchedThreadhold4PPT;
            bool isLeftHandStretched = (leftshoulder.X - leftHand.X) > ArmStretchedThreadhold4PPT;

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
            startKinect();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            isWindowsClosing = true;
            stopKinect(kinect); 
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
            System.Threading.Thread.Sleep(4000);

            // 获取Kinect音频对象
            KinectAudioSource audiosource = kinect.AudioSource;
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
            var voicecommands = new Choices();
            voicecommands.Add("okay");
            voicecommands.Add("thank you");

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
                string city = e.Result.Text.ToLower();
                if (city == "okay")
                {
                    KeyboardToolkit.Keyboard.Type(Key.F5);
                    labelstart.Visibility = System.Windows.Visibility.Hidden;
                }
                else if (city == "thank you")
                {
                    KeyboardToolkit.Keyboard.Type(Key.Escape);
                    labelesc.Visibility = System.Windows.Visibility.Hidden;
                }
            }
        }

    }
}
