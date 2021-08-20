using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Emgu.CV;
using Emgu.CV.Structure;
using Emgu.CV.CvEnum;
using System.IO;
using System.Diagnostics;
using ClosedXML.Excel;
using System.Data;
using DocumentFormat.OpenXml.Spreadsheet;
namespace MultiFaceRec
{
    public partial class FrmPrincipal : Form
    {
        Image<Bgr, Byte> currentFrame;
        Capture grabber;
        HaarCascade face;
        HaarCascade eye;
        DataTable dt = new DataTable();
        MCvFont font = new MCvFont(FONT.CV_FONT_HERSHEY_TRIPLEX, 0.5d, 0.5d);
        Image<Gray, byte> result, TrainedFace = null;
        Image<Gray, byte> gray = null;
        List<Image<Gray, byte>> trainingImages = new List<Image<Gray, byte>>();
        List<string> labels= new List<string>();
        List<string> NamePersons = new List<string>();
        int ContTrain, NumLabels, t;
        string name, names = null;
        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            int kt = trainingImages.ToArray().Length + 1;
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                string fileName;
                fileName = dlg.FileName;
                this.imageBox1.Image = new Image<Bgr, Byte>(fileName);
               Image bitmap = new Bitmap(fileName);
                 int ktten = fileName.Length;
               int vtbd = fileName.LastIndexOf(@"\");
                 int vtkt = fileName.LastIndexOf(".");
                string t = fileName.Substring(0,vtkt - 1);
               string Ten = t.Substring(vtbd + 1);
               string LoadFaces;
                trainingImages.Add(new Image<Gray, byte>(fileName));
                labels.Add(Ten);
                File.WriteAllText(Application.StartupPath + "/TrainedFaces/TrainedLabels.txt", trainingImages.ToArray().Length.ToString() + "%");
                for(int i = 1;i < trainingImages.ToArray().Length + 1; ++i)
                {
                    trainingImages.ToArray()[i - 1].Save(Application.StartupPath + "/TrainedFaces/face" + i + ".bmp");
                    File.AppendAllText(Application.StartupPath + "/TrainedFaces/TrainedLabels.txt", labels.ToArray()[i - 1] + "%");
                }
                MessageBox.Show(Ten);
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "All file | *.*;";
            dlg.Multiselect = true; //Chỗ này nè
            int kt = trainingImages.ToArray().Length + 1;
            if (dlg.ShowDialog(this) == DialogResult.OK)
            {
                foreach (string str in dlg.FileNames)
                {
                    int vtbd = str.LastIndexOf(@"\");
                    int vtkt = str.LastIndexOf(".");
                    string t = str.Substring(0, vtkt - 1);
                    string Ten = t.Substring(vtbd + 1);
                    trainingImages.Add(new Image<Gray, byte>(str));
                    labels.Add(Ten);
                }
                 File.WriteAllText(Application.StartupPath + "/TrainedFaces/TrainedLabels.txt", trainingImages.ToArray().Length.ToString() + "%");
                for(int i = 1;i < trainingImages.ToArray().Length + 1; ++i)
                {
                    trainingImages.ToArray()[i - 1].Save(Application.StartupPath + "/TrainedFaces/face" + i + ".bmp");
                    File.AppendAllText(Application.StartupPath + "/TrainedFaces/TrainedLabels.txt", labels.ToArray()[i - 1] + "%");
                }
            }
        }
        private void button5_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog(){Filter = "Excel Workbook|*.xlsx"})
            {
               
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        using (XLWorkbook workbook = new XLWorkbook())
                        {
                            workbook.Worksheets.Add(dt, DateTime.Now.ToString("dd-MM-yyyy"));             
                            workbook.SaveAs(sfd.FileName);
                        }
                        MessageBox.Show("Xuất file thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch(Exception ex)
                    {
                        MessageBox.Show(ex.ToString(),"Thông báo",MessageBoxButtons.OK,MessageBoxIcon.Error);
                    }
                }
            }
        }
        public FrmPrincipal()
        {
            InitializeComponent();
            face = new HaarCascade("haarcascade_frontalface_default.xml");
            dt.Columns.Add("STT", typeof(int));
            dt.Columns.Add("Tên", typeof(String));
            dt.Columns.Add("Thời gian điểm danh", typeof(String));
            try
            {
                string Labelsinfo = File.ReadAllText(Application.StartupPath + "/TrainedFaces/TrainedLabels.txt");
                string[] Labels = Labelsinfo.Split('%');
                NumLabels = Convert.ToInt16(Labels[0]);
                ContTrain = NumLabels;
                string LoadFaces;
                for (int tf = 1; tf < NumLabels+1; tf++)
                {
                    LoadFaces = "face" + tf + ".bmp";
                    trainingImages.Add(new Image<Gray, byte>(Application.StartupPath + "/TrainedFaces/" + LoadFaces));
                    labels.Add(Labels[tf]);
                }
            }
            catch(Exception e)
            {
                MessageBox.Show("Không có gì trong cơ sở dữ liệu nhị phân, vui lòng thêm ít nhất một khuôn mặt(Chỉ cần đào tạo nguyên mẫu bằng nút Thêm khuôn mặt).", "Tải các khuôn mặt được đào tạo", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            grabber = new Capture();
            grabber.QueryFrame();
            Application.Idle += new EventHandler(FrameGrabber);
            button1.Enabled = false;
        }
        private void button2_Click(object sender, System.EventArgs e)
        {
            try
            { 
                ContTrain = ContTrain + 1; 
                gray = grabber.QueryGrayFrame().Resize(320, 240, Emgu.CV.CvEnum.INTER.CV_INTER_CUBIC); 
                MCvAvgComp[][] facesDetected = gray.DetectHaarCascade(
                face,
                1.2,
                10,
                Emgu.CV.CvEnum.HAAR_DETECTION_TYPE.DO_CANNY_PRUNING,
                new Size(20, 20));
                foreach (MCvAvgComp f in facesDetected[0])
                {
                    TrainedFace = currentFrame.Copy(f.rect).Convert<Gray, byte>();
                    break;
                }   
                TrainedFace = result.Resize(100, 100, Emgu.CV.CvEnum.INTER.CV_INTER_CUBIC);
                trainingImages.Add(TrainedFace);
                labels.Add(textBox1.Text);        
                imageBox1.Image = TrainedFace;
                File.WriteAllText(Application.StartupPath + "/TrainedFaces/TrainedLabels.txt", trainingImages.ToArray().Length.ToString() + "%");       
                for (int i = 1; i < trainingImages.ToArray().Length + 1; i++)
                {
                    MessageBox.Show(labels.ToArray()[i - 1]);
                    trainingImages.ToArray()[i - 1].Save(Application.StartupPath + "/TrainedFaces/face" + i + ".bmp");
                    File.AppendAllText(Application.StartupPath + "/TrainedFaces/TrainedLabels.txt", labels.ToArray()[i - 1] + "%");
                }
                MessageBox.Show(textBox1.Text + "´s khuôn mặt được phát hiện và thêm vào :)", "Đào tạo OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch
            {
                MessageBox.Show("Bật tính năng nhận diện khuôn mặt trước tiên", "Đào tạo thất bại", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }


        void FrameGrabber(object sender, EventArgs e)
        {
            label3.Text = "0";
            NamePersons.Add("");
            currentFrame = grabber.QueryFrame().Resize(320, 240, Emgu.CV.CvEnum.INTER.CV_INTER_CUBIC);
            gray = currentFrame.Convert<Gray, Byte>();      
                    MCvAvgComp[][] facesDetected = gray.DetectHaarCascade(
                  face,
                  1.2,
                  10,
                  Emgu.CV.CvEnum.HAAR_DETECTION_TYPE.DO_CANNY_PRUNING,
                  new Size(20, 20));
                    foreach (MCvAvgComp f in facesDetected[0])
                    {
                        t = t + 1;
                result = currentFrame.Copy(f.rect).Convert<Gray, byte>().Resize(100, 100, Emgu.CV.CvEnum.INTER.CV_INTER_CUBIC);
                        currentFrame.Draw(f.rect, new Bgr(System.Drawing.Color.Black), 2);
                        if (trainingImages.ToArray().Length != 0)
                        {
                           MCvTermCriteria termCrit = new MCvTermCriteria(ContTrain, 0.001);
                           EigenObjectRecognizer recognizer = new EigenObjectRecognizer(
                           trainingImages.ToArray(),
                           labels.ToArray(),
                           3000,
                           ref termCrit);

                           name = recognizer.Recognize(result);
                    currentFrame.Draw(name, ref font, new Point(f.rect.X - 2, f.rect.Y - 2), new Bgr(System.Drawing.Color.LightGreen));
                        }
                            NamePersons[t-1] = name;
                            NamePersons.Add("");   
                        label3.Text = facesDetected[0].Length.ToString();
                    }
                        t = 0;                  
                    for (int nnn = 0; nnn < facesDetected[0].Length; nnn++)
                    {
                        names = names + NamePersons[nnn] + ", ";
                    }
            imageBoxFrameGrabber.Image = currentFrame;
            label4.Text = names;
            if (name.Length != 0)
            {
                dt.Rows.Add("123", name, DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss"));
            }
            names = "";    
                    NamePersons.Clear();
                }
    }
}