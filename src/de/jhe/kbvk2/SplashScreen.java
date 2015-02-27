/**
 * 
 */
package de.jhe.kbvk2;

import android.app.Activity;
import android.content.Intent;
import android.os.Bundle;
import android.view.MotionEvent;

/**
 * @author Jens Helweg
 *
 */
public class SplashScreen extends Activity {
	
	protected boolean active = true;
	protected int splashTime = 2000; // time to display the splash screen in ms
	protected Activity activity;
	
	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
	    setContentView(R.layout.splash);
	    activity = this;
	    
	    // thread for displaying the SplashScreen
	    Thread splashTread = new Thread() {
	        @Override
	        public void run() {
	            try {
	                int waited = 0;
	                while(active && (waited < splashTime)) {
	                    sleep(100);
	                    if(active) {
	                        waited += 100;
	                    }
	                }
	            } catch(InterruptedException e) {
	                // do nothing
	            } finally {
	                finish();
	                Intent startMenuFlipper = new Intent(activity, MenuFlipper.class);
	        		startActivity(startMenuFlipper);
	            }
	        }
	    };
	    splashTread.start();
	}
	
	@Override
	public boolean onTouchEvent(MotionEvent event) {
	    if (event.getAction() == MotionEvent.ACTION_DOWN) {
	        active = false;
	    }
	    return true;
	}
}
