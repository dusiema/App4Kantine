/**
 * 
 */
package de.jhe.kbvk2.gesture;

import android.util.Log;
import android.view.GestureDetector.SimpleOnGestureListener;
import android.view.MotionEvent;
import android.view.animation.AnimationUtils;
import android.widget.ViewFlipper;
import de.jhe.kbvk2.MenuFlipper;
import de.jhe.kbvk2.R;

/**
 * @author Jens Helweg
 * 
 */
public class MyGestureDetector extends SimpleOnGestureListener {
	private static final int SWIPE_MIN_DISTANCE = 120;
	private static final int SWIPE_MAX_OFF_PATH = 250;
	private static final int SWIPE_THRESHOLD_VELOCITY = 200;
	private MenuFlipper menuFlipper;
	private ViewFlipper flipper;

	public MyGestureDetector(MenuFlipper menuFlipper) {
		super();
		this.menuFlipper = menuFlipper;
	}

	@Override
	public boolean onFling(MotionEvent e1, MotionEvent e2, float velocityX,
			float velocityY) {
		try {
			if (Math.abs(e1.getY() - e2.getY()) > SWIPE_MAX_OFF_PATH)
				return false;
			// right to left swipe
			if (e1.getX() - e2.getX() > SWIPE_MIN_DISTANCE
					&& Math.abs(velocityX) > SWIPE_THRESHOLD_VELOCITY) {
				
				getFlipper().setOutAnimation(AnimationUtils.loadAnimation(
						menuFlipper, R.anim.anim_translate_left_out));
				getFlipper().setInAnimation(AnimationUtils.loadAnimation(
						menuFlipper, R.anim.anim_translate_left_in));
				
				getFlipper().showNext();

				menuFlipper.flipForward();
				
			} else if (e2.getX() - e1.getX() > SWIPE_MIN_DISTANCE
					&& Math.abs(velocityX) > SWIPE_THRESHOLD_VELOCITY) {

				getFlipper().setOutAnimation(AnimationUtils.loadAnimation(
						menuFlipper, R.anim.anim_translate_right_out));
				getFlipper().setInAnimation(AnimationUtils.loadAnimation(
						menuFlipper, R.anim.anim_translate_right_in));
				
				getFlipper().showPrevious();

				menuFlipper.flipBackward();

			}
		} catch (Exception e) {
			Log.d(this.getClass().toString(), "Error in onFling() method!");
		}
		return false;
	}
	
	public ViewFlipper getFlipper() {
		if (flipper == null) {
			flipper = (ViewFlipper) menuFlipper
					.findViewById(R.id.menuFlipper);
		}
		return flipper;
	}
	

}
