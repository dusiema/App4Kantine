/**
 * 
 */
package de.jhe.kbvk2;

import android.app.Activity;
import android.app.ProgressDialog;
import android.content.Intent;
import android.os.Bundle;
import android.view.View;
import android.widget.Button;
import de.jhe.kbvk2.async.DownloadFileAsync;

/**
 * @author jhelweg
 * 
 */
public class MenuChooser extends Activity {

	private static final String MENU_URL = "http://a.ndroi.de/android/speiseplan.php";

	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.menu_chooser);
		
		if (getFilesDir().list().length > 0) {
			Button button = (Button) findViewById(R.id.showMenuFile);
			button.setText(getFilesDir().list()[0]);
		}

	}

	public void downloadNewMenu(View view) {
		ProgressDialog progressDialog = new ProgressDialog(this);
		progressDialog.setProgressStyle(ProgressDialog.STYLE_HORIZONTAL);
		DownloadFileAsync downloadFileAsync = new DownloadFileAsync(this, progressDialog);
		downloadFileAsync.execute(MENU_URL);
		progressDialog.setTitle("Download");
		progressDialog.setMessage("Lade Speiseplan...");
		progressDialog.show();
	}
	
	public void startMenuFlipper(View view) {
		Intent startMenuFlipper = new Intent(this, MenuFlipper.class);
		startActivity(startMenuFlipper);
	}

}
