package com.aspose.slides.helloworld;

import java.io.FileInputStream;

import com.aspose.slides.AutoShapeEx;
import com.aspose.slides.FillTypeEx;
import com.aspose.slides.Presentation;
import com.aspose.slides.PresentationEx;
import com.aspose.slides.SaveFormat;
import com.aspose.slides.ShapeTypeEx;
import com.aspose.slides.Slide;
import com.aspose.slides.SlideEx;

import android.os.Bundle;
import android.os.Environment;
import android.app.Activity;
import android.graphics.Color;
import android.util.Log;
import android.view.Menu;
import android.widget.TextView;

//import com.aspose.slides.*;
//import com.example.mytestapp.R;

public class MainActivity extends Activity {

	private static final String TAG = "Aspose.Slides.Test";
	static  String path = Environment.getExternalStorageDirectory().getPath()+"//";


	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_main);
      //  final TextView tx = (TextView) findViewById(R.id.textView1);

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
				
				   //Generating PPT 
				   HelloWorld_PPT();
				   
				   //Generating PPTX 
				   HelloWorldPPTX();

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

	 private static void HelloWorld_PPT()
	    {
	    	try{
	    		
	    	

			       Log.i(TAG,"Loading Pres..");
			       
			    //Instantiate a Presentation object that represents a PPT file
		    	Presentation pres = new Presentation();
		
		        Log.i(TAG,"Pres loaded....");
		        		        
		    	//Adding an empty slide to the presentation and getting the reference of
		    	//that empty slide
		    	Slide slide = pres.addEmptySlide();
		
		
		    	//Adding a rectangle (X=2400, Y=1800, Width=1000 & Height=500) to the slide
		    	com.aspose.slides.Rectangle rect = slide.getShapes().addRectangle(2400, 1800, 1000, 500);
		
		
		    	//Hiding the lines of rectangle
		    	rect.getLineFormat().setShowLines(false);
		
		
		    	//Adding a text frame to the rectangle with "Hello World" as a default text
		    	rect.addTextFrame("Hello World, I am alive in Android... :)");
		
		
		    	//Removing the first slide of the presentation which is always added by
		    	//Aspose.Slides for Java by default while creating the presentation
		    	pres.getSlides().removeAt(0);
		
		
		    	//Writing the presentation as a PPT file
		    	pres.write(path+"HelloAndroid.ppt");
	    	
		    	Log.i(TAG, "CreateHelloWorld_PPT Sucess...");
		    	   
	    	}
	    	catch(Exception e)
	    	{
	    		//Log.e(TAG, e.getMessage());
		    	Log.i(TAG, "CreateHelloWorld_PPT Failed...");
		    	Log.e(TAG, e.getMessage());
	    	}
	    	
	    }

		//Working with PPTX
		private static void  HelloWorldPPTX()
		{

			try{

		        Log.i(TAG,"Loading PPTX..");

				//Instantiate PresentationEx
				PresentationEx pres = new PresentationEx();

				//Get the first slide
				SlideEx sld = pres.getSlides().get_Item(0);

				//Add an AutoShape of Rectangle type
				int idx = sld.getShapes().addAutoShape(ShapeTypeEx.Rectangle, 150, 75,150 , 50);
				AutoShapeEx ashp = (AutoShapeEx)sld.getShapes().get_Item(idx);

				//Add TextFrame to the Rectangle
				ashp.addTextFrame("Hello World I am alive in Android");

				//Change the text color to Black (which is White by default)
				ashp.getTextFrame().getParagraphs().get_Item(0).getPortions().get_Item(0).getFillFormat().setFillType(FillTypeEx.Solid);
				ashp.getTextFrame().getParagraphs().get_Item(0).getPortions().get_Item(0).getFillFormat().getSolidFillColor().setColor(Color.BLACK);

				//Change the line color of the rectangle to White
				ashp.getShapeStyle().getLineColor().setColor(Color.WHITE);

				//Remove any fill formatting in the shape
				ashp.getFillFormat().setFillType(FillTypeEx.NoFill);
				            
			
				 Log.i(TAG,"Saving PPTX..");

				//Write the presentation to disk
				pres.write(path+"HelloAndroid.pptx");
				
				Log.i(TAG,"Success: HelloWorldPPTX..");
				
				}
			catch(Exception e)
			{
					Log.e(TAG, e.getMessage());
			       Log.i(TAG,"Error Loading: HelloWorldPPTX..");
			
			}
			
		}	


}
