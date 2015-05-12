package com.example.addingshapes;

import java.io.File;
import java.io.FileInputStream;

import com.aspose.slides.*;

import android.os.Bundle;
import android.os.Environment;
import android.app.Activity;
import android.graphics.Color;
import android.util.Log;
import android.view.Menu;

public class MainActivity extends Activity {

	private static final String TAG = "Aspose.Slides.Test";
	static  String path = Environment.getExternalStorageDirectory().getPath()+"//";

	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_main);
		
	    try{
				Log.i(TAG,"Start Loading license..");
				
				//Create a License object
				com.aspose.slides.License license = new com.aspose.slides.License();
				
				//Create a stream object containing the license file
				   FileInputStream fstream=new FileInputStream(path+"Aspose.Total.Java.lic");
				
			   //Set the license through the stream object
				   license.setLicense(fstream); 
				
				   fstream.close();
				   
				   Log.i(TAG,"End Loading license..");
				

				   //Working with PPT Shapes
				   AddSimpleLine();
				   AddAdvancedLine();
				   AddSimpleEllipse();
				   AddAdvancedEllipse();
				   AddSimpleRectangle();
				   AddAdvancedRectangle();
				   AddPictureFrame();
				   AddingAudioFrametoSlide();
				   AddingVideoFrametoSlide();
				   AddingOleFrametoSlide();

				   
				   //Working with PPTX Shapes
				   AddSimpleLine_PPTX();
				   AddAdvancedLine_PPTX();
				   AddSimpleEllipse_PPTX();
				   AddAdvancedEllipse_PPTX();
				   AddSimpleRectangle_PPTX();
				   AddAdvancedRectangle_PPTX();
				   AddPictureFrame_PPTX();
				   AddingAudioFrametoSlide_PPTX();
				   AddingVideoFrametoSlide_PPTX();
				   AddingEmbeddedVideoFrametoSlide_PPTX();
				   AddingOleFrametoSlide_PPTX();
				   
				   
	}

	catch(Exception e)
	{
	e.printStackTrace();
	}

	}

	@Override
	public boolean onCreateOptionsMenu(Menu menu) {
		// Inflate the menu; this adds items to the action bar if it is present.
		getMenuInflater().inflate(R.menu.main, menu);
		return true;
	}
	
	//Working with PPT shapes
	private void AddSimpleLine()
	{
		
		try{
			
			//Instantiate a Presentation object that represents a PPT file
			Presentation pres = new Presentation();

			//Accessing a slide using its slide position
			Slide slide = pres.getSlideByPosition(1);

			
			 android.graphics.Point point1 = new android.graphics.Point(1000,2000);
			 android.graphics.Point point2 = new android.graphics.Point(4700,2000);
			 
			//Adding a line shape into the slide with its start and end points
			slide.getShapes().addLine(point1, point2);

			//Writing the presentation as a PPT file
			pres.write(path+"AddSimpleLine.ppt");
			//Working with Shapes
   	        
			Log.i(TAG,"Success: AddSimpleLine..");
 		
			}
		catch(Exception e)
		{
 			Log.e(TAG, e.getMessage());
   	        Log.i(TAG,"Error Loading: AddSimpleLine..");
 
		}
		
	}
	
	private void AddAdvancedLine()
	{
		
		try{
			
			//Instantiate a Presentation object that represents a PPT file
			Presentation pres = new Presentation();

			//Accessing a slide using its slide position
			Slide slide = pres.getSlideByPosition(1);


			android.graphics.Point point1 = new android.graphics.Point(1000,2000);
			android.graphics.Point point2 = new android.graphics.Point(4700,2000);

			//Adding a line shape into the slide with its start and end points
			Line shape=slide.getShapes().addLine(point1, point2);


			//Setting the fore color of the line to blue
			shape.getLineFormat().setForeColor(android.graphics.Color.BLUE);


			//Setting the width of the line to 10
			shape.getLineFormat().setWidth((byte)10);


			//Setting the line style to thick between thin
			shape.getLineFormat().setStyle(LineStyle.ThickBetweenThin);


			//Setting the dash style of the line to dash
			shape.getLineFormat().setDashStyle(LineDashStyle.Dash);


			//Setting the style of the starting point of the line to oval
			shape.getLineFormat().setBeginArrowheadStyle(LineArrowheadStyle.Oval);


			//Setting the style of the ending point of the line to open
			shape.getLineFormat().setEndArrowheadStyle(LineArrowheadStyle.Open);

			
			//Writing the presentation as a PPT file
			pres.write(path+"AdvancedLine.ppt");
			//Working with Shapes
   	        
			Log.i(TAG,"Success: AdvancedLine..");
 		
			}
		catch(Exception e)
		{
 			Log.e(TAG, e.getMessage());
   	        Log.i(TAG,"Error Loading: AddAdvancedLine..");
 
		}
		
	}
	
	private void AddSimpleEllipse()
	{
		
		try{
			
			//Instantiate a Presentation object that represents a PPT file
			Presentation pres = new Presentation();

			//Accessing a slide using its slide position
			Slide slide = pres.getSlideByPosition(1);


			//Adding an ellipse shape into the slide by defining its X,Y postion, width
			//and height
			slide.getShapes().addEllipse(2300, 1200, 1000, 2000);
			//Writing the presentation as a PPT file

			pres.write(path+"AddSimpleEllipse.ppt");
			//Working with Shapes
   	        
			Log.i(TAG,"Success: AddSimpleEllipse..");
 		
			}
		catch(Exception e)
		{
 			Log.e(TAG, e.getMessage());
   	        Log.i(TAG,"Error Loading: AddSimpleEllipse..");
 
		}
		
	}
	
	private void AddAdvancedEllipse()
	{
		
		try{
			
			//Instantiate a Presentation object that represents a PPT file
			Presentation pres = new Presentation();

			//Accessing a slide using its slide position
			Slide slide = pres.getSlideByPosition(1);


			//Setting the fore color of the line to blue


			//Adding an ellipse shape into the slide by defining its X,Y postion, width
			//and height
			Shape shape = slide.getShapes().addEllipse(2300, 1200, 1000, 2000);


			//Setting the fill type of the ellipse to gradient
			shape.getFillFormat().setType(FillType.Gradient);


			//Setting the color type of the gradient to two colors
			shape.getFillFormat().setGradientColorType(GradientColorType.TwoColors);


			//Setting the background color of the ellipse to red
			shape.getFillFormat().setBackColor(android.graphics.Color.RED);


			//Setting the foreground color of the ellipse to blue
			shape.getFillFormat().setForeColor(android.graphics.Color.BLUE);


			//Setting the fill angle of the gradient to 90
			shape.getFillFormat().setGradientFillAngle((byte)90);


			//Setting the rotation of the ellipse to 45 degrees
			shape.setRotation((byte)45);


			//Setting the foreground color of the ellipse lines
			shape.getLineFormat().setForeColor(android.graphics.Color.YELLOW);


			
			//Writing the presentation as a PPT file
			pres.write(path+"AdvancedEllipse.ppt");
			//Working with Shapes
   	        
			Log.i(TAG,"Success: AdvancedEllipse..");
 		
			}
		catch(Exception e)
		{
 			Log.e(TAG, e.getMessage());
   	        Log.i(TAG,"Error Loading: AdvancedEllipse..");
 
		}
		
	}
	
	private void AddSimpleRectangle()
	{
		
		try{
			
			//Instantiate a Presentation object that represents a PPT file
			Presentation pres = new Presentation();

			//Accessing a slide using its slide position
			Slide slide = pres.getSlideByPosition(1);


			//Adding a rectangle shape into the slide by defining its X,Y postion, width
			//and height
			slide.getShapes().addRectangle(1400, 1100, 3000, 2000);

			
			//Writing the presentation as a PPT file

			pres.write(path+"AddSimpleRectangle.ppt");
			//Working with Shapes
   	        
			Log.i(TAG,"Success: AddSimpleRectangle..");
 		
			}
		catch(Exception e)
		{
 			Log.e(TAG, e.getMessage());
   	        Log.i(TAG,"Error Loading: AddSimpleRectangle..");
 
		}
		
	}
	
	private void AddAdvancedRectangle()
	{
		
		try{
			
			//Instantiate a Presentation object that represents a PPT file
			Presentation pres = new Presentation();

			//Accessing a slide using its slide position
			Slide slide = pres.getSlideByPosition(1);


			//Adding a rectangle shape into the slide by defining its X,Y postion, width
			//and height
			Shape shape = slide.getShapes().addRectangle(1400, 1100, 3000, 2000);


			//Setting the fill type of the rectangle to pattern
			shape.getFillFormat().setType(FillType.Pattern);


			//Setting the pattern style to sphere
			shape.getFillFormat().setPatternStyle(PatternStyle.Sphere);


			//Setting the background color of the rectangle to light gray
			shape.getFillFormat().setBackColor(android.graphics.Color.GRAY);


			//Setting the foreground color of the rectangle to yellow
			shape.getFillFormat().setForeColor(android.graphics.Color.YELLOW);


			//Setting the foreground color of the rectangle lines to blue
			shape.getLineFormat().setForeColor(android.graphics.Color.BLUE);


			//Setting the width of the rectangle lines in points
			shape.getLineFormat().setWidth((byte)10);


			//Setting the line style of the rectangle to thick thin
			shape.getLineFormat().setStyle(LineStyle.ThickThin);


			
			//Writing the presentation as a PPT file
			pres.write(path+"AddAdvancedRectangle.ppt");
			//Working with Shapes
   	        
			Log.i(TAG,"Success: AddAdvancedRectangle..");
 		
			}
		catch(Exception e)
		{
 			Log.e(TAG, e.getMessage());
   	        Log.i(TAG,"Error Loading: AddAdvancedRectangle..");
 
		}
		
	}

	private void AddPictureFrame()
	{
		
		try{
			
			//Instantiate a Presentation object that represents a PPT file
			Presentation pres = new Presentation();

			//Accessing a slide using its slide position
			Slide slide = pres.getSlideByPosition(1);

			//Creating a Buffered Image object to hold the image file
			//Setting the fill type of the background to picture
			slide.getBackground().getFillFormat().setType(FillType.Picture);

			//Creating a Buffered Image object to hold the image file
			android.graphics.Bitmap image= 	android.graphics.BitmapFactory.decodeFile(path+"Desert.jpg");

			//Creating a picture object that will be used as a slide background
			com.aspose.slides.Picture pic = new com.aspose.slides.Picture(pres,image);


			//Adding the picture object to pictures collection of the presentation
			//After the picture object is added, the picture is given a uniqe picture Id
			int picId = pres.getPictures().add(pic);


			//Calculating picture width and height
			int pictureWidth = pres.getPictures().get_Item(picId - 1).getImage().getWidth() * 4;
			int pictureHeight = pres.getPictures().get_Item(picId - 1).getImage().getHeight() * 4;


			//Calculating slide width and height
			int slideWidth = slide.getBackground().getWidth();
			int slideHeight = slide.getBackground().getHeight();


			//Calculating the width and height of picture frame

			int pictureFrameWidth = (int)(slideWidth / 2 - pictureWidth / 2);
			int pictureFrameHeight = (int)(slideHeight / 2 - pictureHeight / 2);

			//Adding picture frame to the slide
			slide.getShapes().addPictureFrame(picId, pictureFrameWidth, pictureFrameHeight,
			                                        pictureWidth, pictureHeight);
			//Writing the presentation as a PPT file
			pres.write(path+"AddPictureFrame.ppt");
			//Working with Shapes
   	        
			Log.i(TAG,"Success: AddPictureFrame..");
 		
			}
		catch(Exception e)
		{
 			Log.e(TAG, e.getMessage());
   	        Log.i(TAG,"Error Loading: AddPictureFrame..");
 
		}
		
	}
	
	private void AddingAudioFrametoSlide()
	{
		
		try{
			
			//Instantiate a Presentation object that represents a PPT file
			Presentation pres = new Presentation();

			//Accessing a slide using its slide position
			Slide slide = pres.getSlideByPosition(1);

			//Opening the audio file as a stream
			FileInputStream fstr = new FileInputStream(new File(path+"Kalimba.mp3"));

			//Adding the embedded audio frame into the slide
			slide.getShapes().addAudioFrameEmbedded(2600, 1800, 500, 500, fstr);

			                                        
			//Writing the presentation as a PPT file
			pres.write(path+"AddingAudioFrametoSlide.ppt");
			//Working with Shapes
   	        
			Log.i(TAG,"Success: AddingAudioFrametoSlide..");
 		
			}
		catch(Exception e)
		{
 			Log.e(TAG, e.getMessage());
   	        Log.i(TAG,"Error Loading: AddingAudioFrametoSlide..");
 
		}
		
	}
	private void AddingVideoFrametoSlide()
	{
		
		try{
			
			//Instantiate a Presentation object that represents a PPT file
			Presentation pres = new Presentation();

			//Accessing a slide using its slide position
			Slide slide = pres.getSlideByPosition(1);

			//Adding the video frame into the slide
			slide.getShapes().addVideoFrame(1900, 1100, 2000, 2000, path+"Wildlife.wmv");

			                                        
			//Writing the presentation as a PPT file
			pres.write(path+"AddingVideoFrametoSlide.ppt");
			//Working with Shapes
   	        
			Log.i(TAG,"Success: AddingVideoFrametoSlide..");
 		
			}
		catch(Exception e)
		{
 			Log.e(TAG, e.getMessage());
   	        Log.i(TAG,"Error Loading: AddingVideoFrametoSlide..");
 
		}
		
	}

	private void AddingOleFrametoSlide()
	{
		
		try{
			
			//Instantiate a Presentation object that represents a PPT file
			Presentation pres = new Presentation();

			//Accessing a slide using its slide position
			Slide slide = pres.getSlideByPosition(1);

			//Reading excel chart from the excel file and save as an array of bytes
			File file=new File(path+"Test.xls");

			int length=(int)file.length();

			FileInputStream fstro = new FileInputStream(file);

			byte[] b = new byte[length];
			fstro.read(b, 0, length);


			//Inserting the excel chart as new OleObjectFrame to a slide
			OleObjectFrame oof = slide.getShapes().addOleObjectFrame(0, 0, (int)pres.getSlideSize().width(),(int)pres.getSlideSize().height(), "Excel.Sheet.8", b);
			                                        
			//Writing the presentation as a PPT file
			pres.write(path+"AddingOleFrametoSlide.ppt");
			//Working with Shapes
   	        
			Log.i(TAG,"Success: AddingOleFrametoSlide..");
 		
			}
		catch(Exception e)
		{
 			Log.e(TAG, e.getMessage());
   	        Log.i(TAG,"Error Loading: AddingOleFrametoSlide..");
 
		}
		
	}
	
	//Working with PPTX shapes
	private void AddSimpleLine_PPTX()
	{
		
		try{
			
			//Instantiate a Presentation object that represents a PPT file
			PresentationEx pres = new PresentationEx();

			//Get the first slide
			SlideEx sld = pres.getSlides().get_Item(0);

			//Add an autoshape of type line
			int idx = sld.getShapes().addAutoShape(ShapeTypeEx.Line, 50, 150, 300, 0);
			            
			
			//Writing the presentation as a PPT file
			pres.write(path+"AddSimpleLine_PPTX.pptx");
			//Working with Shapes
   	        
			Log.i(TAG,"Success: AddSimpleLine_PPTX..");
 		
			}
		catch(Exception e)
		{
 			Log.e(TAG, e.getMessage());
   	        Log.i(TAG,"Error Loading: AddSimpleLine_PPTX..");
 
		}
		
	}
	
	private void AddAdvancedLine_PPTX()
	{
		
		try{
			
			//Instantiate a Presentation object that represents a PPTX file
			PresentationEx pres = new PresentationEx();
			
			//Get the first slide
			SlideEx sld = pres.getSlides().get_Item(0);
			
			//Add an autoshape of type line
			int idx = sld.getShapes().addAutoShape(ShapeTypeEx.Line, 50, 150, 300, 0);
			ShapeEx shp = sld.getShapes().get_Item(idx);
			     
			//Apply some formatting on the line
			shp.getLineFormat().setStyle(LineStyleEx.ThickBetweenThin);
			shp.getLineFormat().setWidth((byte)10);
			            
			shp.getLineFormat().setDashStyle(LineDashStyleEx.DashDot);
			
			shp.getLineFormat().setBeginArrowheadLength(LineArrowheadLengthEx.Short);
			shp.getLineFormat().setBeginArrowheadStyle(LineArrowheadStyleEx.Oval);
			
			shp.getLineFormat().setEndArrowheadLength(LineArrowheadLengthEx.Long);
			shp.getLineFormat().setEndArrowheadStyle(LineArrowheadStyleEx.Triangle);
			
			shp.getLineFormat().getFillFormat().setFillType(FillTypeEx.Solid);
			shp.getLineFormat().getFillFormat().getSolidFillColor().setColor(Color.RED);
			
			//Writing the presentation as a PPT file
			pres.write(path+"AdvancedLine_PPTX.pptx");
			//Working with Shapes
   	        
			Log.i(TAG,"Success: AdvancedLine_PPTX..");
 		
			}
		catch(Exception e)
		{
 			Log.e(TAG, e.getMessage());
   	        Log.i(TAG,"Error Loading: AddAdvancedLine_PPTX..");
 
		}
		
	}
	
	private void AddSimpleEllipse_PPTX()
	{
		
		try{
			
			//Instantiate a Presentation object that represents a PPT file
			PresentationEx pres = new PresentationEx();

			//Get the first slide
			SlideEx sld = pres.getSlides().get_Item(0);

			//Add an autoshape of type line
			int idx = sld.getShapes().addAutoShape(ShapeTypeEx.Ellipse, 50, 150, 300, 400);
			ShapeEx shp = sld.getShapes().get_Item(idx);

			//Writing the presentation as a PPT file

			pres.write(path+"AddSimpleEllipse_PPTX.pptx");
			//Working with Shapes
   	        
			Log.i(TAG,"Success: AddSimpleEllipse_PPTX..");
 		
			}
		catch(Exception e)
		{
 			Log.e(TAG, e.getMessage());
   	        Log.i(TAG,"Error Loading: AddSimpleEllipse_PPTX..");
 
		}
		
	}
	
	private void AddAdvancedEllipse_PPTX()
	{
		
		try{
			
			//Instantiate a Presentation object that represents a PPT file
			PresentationEx pres = new PresentationEx();

			//Get the first slide
			SlideEx sld = pres.getSlides().get_Item(0);

			//Add autoshape of ellipse type
			int idx = sld.getShapes().addAutoShape(ShapeTypeEx.Ellipse, 50, 150, 75, 150);
			ShapeEx shp = sld.getShapes().get_Item(idx);

			//Apply some gradiant formatting to ellipse shape
			shp.getFillFormat().setFillType( FillTypeEx.Gradient);
			shp.getFillFormat().getGradientFormat().setGradientShape(GradientShapeEx.Linear);

			//Set the Gradient Direction
			shp.getFillFormat().getGradientFormat().setGradientDirection(GradientDirectionEx.FromCorner2);

			//Add two Gradiant Stops
			shp.getFillFormat().getGradientFormat().getGradientStops().add((float)1.0, PresetColorEx.Purple);
			shp.getFillFormat().getGradientFormat().getGradientStops().add((float)0, PresetColorEx.Red);
			
			//Writing the presentation as a PPT file
			pres.write(path+"AdvancedEllipse_PPTX.pptx");
			//Working with Shapes
   	        
			Log.i(TAG,"Success: AdvancedEllipse_PPTX..");
 		
			}
		catch(Exception e)
		{
 			Log.e(TAG, e.getMessage());
   	        Log.i(TAG,"Error Loading: AdvancedEllipse_PPTX..");
 
		}
		
	}
	
	private void AddSimpleRectangle_PPTX()
	{
		
		try{
			
			//Instantiate a Presentation object that represents a PPT file
			PresentationEx pres = new PresentationEx();

			//Accessing a slide using its slide position
			SlideEx slide = pres.getSlides().get_Item(0);


			//Adding a rectangle shape into the slide by defining its X,Y postion, width
			//and height
			slide.getShapes().addAutoShape(ShapeTypeEx.Rectangle,200, 100, 300, 500);

			
			//Writing the presentation as a PPT file

			pres.write(path+"AddSimpleRectangle_PPTX.pptx");
			//Working with Shapes
   	        
			Log.i(TAG,"Success: AddSimpleRectangle_PPTX..");
 		
			}
		catch(Exception e)
		{
 			Log.e(TAG, e.getMessage());
   	        Log.i(TAG,"Error Loading: AddSimpleRectangle_PPTX..");
 
		}
		
	}
	
	private void AddAdvancedRectangle_PPTX()
	{
		
		try{
			
			//Instantiate PrseetationEx class that represents the PPTX
			PresentationEx pres = new PresentationEx();

			//Get the first slide
			SlideEx sld = pres.getSlides().get_Item(0);

			//Add autoshape of rectangle type
			int idx = sld.getShapes().addAutoShape(ShapeTypeEx.Rectangle, 50, 150, 75, 150);
			ShapeEx shp = sld.getShapes().get_Item(idx);

			//Set the fill type to Pattern
			shp.getFillFormat().setFillType(FillTypeEx.Pattern);

			//Set the pattern style
			shp.getFillFormat().getPatternFormat().setPatternStyle(PatternStyleEx.Trellis);

			//Set the pattern back and fore colors
			shp.getFillFormat().getPatternFormat().getBackColorFormat().setColor(Color.GRAY);
			shp.getFillFormat().getPatternFormat().getForeColorFormat().setColor(Color.YELLOW);

			
			//Writing the presentation as a PPT file
			pres.write(path+"AddAdvancedRectangle_PPTX.pptx");
			//Working with Shapes
   	        
			Log.i(TAG,"Success: AddAdvancedRectangle_PPTX..");
 		
			}
		catch(Exception e)
		{
 			Log.e(TAG, e.getMessage());
   	        Log.i(TAG,"Error Loading: AddAdvancedRectangle_PPTX..");
 
		}
		
	}

	private void AddPictureFrame_PPTX()
	{
		
		try{
			
			//Instantiate PresentationEx class that represents the PPTX
			PresentationEx pres = new PresentationEx();
			
			//Get the first slide
			SlideEx sld = pres.getSlides().get_Item(0);
			
			//Creating a Buffered Image object to hold the image file
			android.graphics.Bitmap image= 	android.graphics.BitmapFactory.decodeFile(path+"Desert.jpg");
			
			ImageEx imgx = pres.getImages().addImage(image);
			
			
			//Add Picture Frame with height and width equivalent of Picture
			sld.getShapes().addPictureFrame(ShapeTypeEx.Rectangle, 50, 150, imgx.getWidth(), imgx.getHeight(),imgx);
			
			
			//Writing the presentation as a PPTX file
			pres.write(path+"AddPictureFrame_PPTX.pptx");
			//Working with Shapes
   	        
			Log.i(TAG,"Success: AddPictureFrame_PPTX..");
 		
			}
		catch(Exception e)
		{
 			Log.e(TAG, e.getMessage());
   	        Log.i(TAG,"Error Loading: AddPictureFrame_PPTX..");
 
		}
		
	}

	private void AddingAudioFrametoSlide_PPTX()
	{
		
		try{
			
			//Instantiate PrseetationEx class that represents the PPTX
			PresentationEx pres = new PresentationEx();
			
			//Get the first slide
			SlideEx sld = pres.getSlides().get_Item(0);
			
			//Opening the audio file as a stream
			FileInputStream fstr = new FileInputStream(new File(path+"Kalimba.mp3"));
			
			//Add Audio Frame 
			AudioFrameEx af = sld.getShapes().addAudioFrameEmbedded(50, 150, 100, 100, fstr);
			            
			//Set Play Mode and Volume of the Audio
			af.setPlayMode(AudioPlayModePresetEx.Auto);
			af.setVolume(AudioVolumeModeEx.Loud);
			
			//Writing the presentation as a PPT file
			pres.write(path+"AddingAudioFrametoSlide_PPTX.pptx");
			//Working with Shapes
   	        
			Log.i(TAG,"Success: AddingAudioFrametoSlide_PPTX..");
 		
			}
		catch(Exception e)
		{
 			Log.e(TAG, e.getMessage());
   	        Log.i(TAG,"Error Loading: AddingAudioFrametoSlide_PPTX..");
 
		}
		
	}
	private void AddingVideoFrametoSlide_PPTX()
	{
		
		try{
			
			//Instantiate PrseetationEx class that represents the PPTX
			PresentationEx pres = new PresentationEx();
			
			//Get the first slide
			SlideEx sld = pres.getSlides().get_Item(0);
			
			//Add Video Frame 
			VideoFrameEx vf = sld.getShapes().addVideoFrame(50, 150, 300, 150,  path+"Wildlife.wmv");
			
			//Set Play Mode and Volume of the Video
			vf.setPlayMode(VideoPlayModePresetEx.Auto);
			vf.setVolume(AudioVolumeModeEx.Loud);
			            
			                                        
			//Writing the presentation as a PPT file
			pres.write(path+"AddingVideoFrametoSlide_PPTX.pptx");
			//Working with Shapes
   	        
			Log.i(TAG,"Success: AddingVideoFrametoSlide_PPTX..");
 		
			}
		catch(Exception e)
		{
 			Log.e(TAG, e.getMessage());
   	        Log.i(TAG,"Error Loading: AddingVideoFrametoSlide_PPTX..");
 
		}
		
	}

	private void AddingEmbeddedVideoFrametoSlide_PPTX()
	{
		
		try{
			
			//Instantiate PrseetationEx class that represents the PPTX
			PresentationEx pres = new PresentationEx();
			
			//Get the first slide
			SlideEx sld = pres.getSlides().get_Item(0);
			
			//Add Video Frame 
			VideoFrameEx vf = sld.getShapes().addVideoFrame(50, 150, 300, 350, null);
			
			//Embedd vide inside presentation
			VideoEx vid = pres.getVideos().addVideo( new FileInputStream(path+"Wildlife.wmv")); // adding VideoEx from stream
			
			//Set video to Video Frame
			vf.setEmbeddedVideo(vid);
			
			
			//Set Play Mode and Volume of the Video
			vf.setPlayMode(VideoPlayModePresetEx.Auto);
			vf.setVolume(AudioVolumeModeEx.Loud);
			     
			//Creating a Buffered Image object to hold the image file
			android.graphics.Bitmap image= 	android.graphics.BitmapFactory.decodeFile(path+"Desert.jpg");
			
			ImageEx imgx = pres.getImages().addImage(image);
			
			
			//Set the fill type to Picture
			vf.getFillFormat().setFillType(FillTypeEx.Picture);
			
			//setting image to vide frame
			vf.setImage(imgx);
			
			//Writing the presentation as a PPTX file
			pres.write(path+"AddingEmbeddedVideoFrametoSlide_PPTX.pptx");
			//Working with Shapes
   	        
			Log.i(TAG,"Success: AddingEmbeddedVideoFrametoSlide_PPTX..");
 		
			}
		catch(Exception e)
		{
 			Log.e(TAG, e.getMessage());
   	        Log.i(TAG,"Error Loading: AddingEmbeddedVideoFrametoSlide_PPTX..");
 
		}
		
	}

	
	private void AddingOleFrametoSlide_PPTX()
	{
		
		try{
			
			
			//Instantiate PrseetationEx class that represents the PPTX
			PresentationEx pres = new PresentationEx();

			//Access the first slide
			SlideEx sld = pres.getSlides().get_Item(0);

			//Load an Excel file to Array of Bytes
			//Reading excel chart from the excel file and save as an array of bytes
			File file=new File(path+"Test.xls");

			int length=(int)file.length();

			FileInputStream fstro = new FileInputStream(file);

			byte[] buf = new byte[length];
			fstro.read(buf, 0, length);

			
			//Add an Ole Object Frame shape
			OleObjectFrameEx oof = sld.getShapes().addOleObjectFrame((float)0,(float) 0, (float)pres.getSlideSize().getSize().width(),(float) pres.getSlideSize().getSize().height(), "Excel.Sheet.8", buf);

			//Writing the presentation as a PPT file
			pres.write(path+"AddingOleFrametoSlide_PPTX.pptx");
			//Working with Shapes
   	        
			Log.i(TAG,"Success: AddingOleFrametoSlide_PPTX..");
 		
			}
		catch(Exception e)
		{
 			Log.e(TAG, e.getMessage());
   	        Log.i(TAG,"Error Loading: AddingOleFrametoSlide_PPTX..");
 
		}
		
	}

}
