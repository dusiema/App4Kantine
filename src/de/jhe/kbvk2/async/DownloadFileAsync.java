/**
 * 
 */
package de.jhe.kbvk2.async;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.net.URL;

import android.app.Activity;
import android.app.ProgressDialog;
import android.os.AsyncTask;
import android.util.Log;
import de.jhe.kbvk2.MenuFlipper;

/**
 * @author Jens Helweg
 * 
 */
public class DownloadFileAsync extends AsyncTask<String, String, String> {
	private Activity activity;
	private ProgressDialog progressDialog;

	public DownloadFileAsync(Activity activity, ProgressDialog progressDialog) {
		this.activity = activity;
		this.progressDialog = progressDialog;
	}

	@Override
	protected String doInBackground(String... aurl) {

		deleteExistingFiles();
		
		try {
			URL url = new URL(aurl[0]);

			HttpURLConnection httpConnection = (HttpURLConnection) url
					.openConnection();
			httpConnection.setRequestMethod("GET");
			httpConnection.setDoOutput(true);
			httpConnection.connect();

			int lenghtOfFile = httpConnection.getContentLength();
			Log.d("ANDRO_ASYNC", "Lenght of file: " + lenghtOfFile);

			String contentDisposition = httpConnection
					.getHeaderField("Content-Disposition");
			
			String fileName = "unkonwn.xls";
			if (contentDisposition != null) {
				fileName = contentDisposition.substring(
						contentDisposition.indexOf('=')+1,
						contentDisposition.indexOf(';',
								contentDisposition.indexOf('=')));
				
			}
			
			InputStream input = httpConnection.getInputStream();

			// TODO: add some checking for the case when external sdcard is not
			// available

			OutputStream output = new FileOutputStream(new File(activity
					.getFilesDir().getPath() + "/" + fileName));

			byte buffer[] = new byte[1024];

			int len1 = 0;
			int bytesWritten = 0;
			while ((len1 = input.read(buffer)) > 0) {
				output.write(buffer, 0, len1);
				bytesWritten += len1;
				onProgressUpdate(String.valueOf(bytesWritten * 100
						/ lenghtOfFile));
			}

			output.flush();
			output.close();
			input.close();

		} catch (Exception e) {
			Log.d("DOWNLOAD_EXCEL", "Download of excel menu file failed!", e);
		}
		return null;

	}

	protected void onProgressUpdate(String... progress) {

		progressDialog.setProgress(Integer.parseInt(progress[0]));
		// progressDialog.incrementProgressBy(Integer.parseInt(progress[0]));
	}

	@Override
	protected void onPostExecute(String unused) {
		progressDialog.cancel();
		if (activity.getFilesDir().list().length > 0) {
//			Button button = (Button) activity.findViewById(R.id.showMenuFile);
//			button.setText(activity.getFilesDir().list()[0]);
			
			if (activity instanceof MenuFlipper) {
				MenuFlipper flipper = (MenuFlipper) activity;
				flipper.loadExcel(flipper.getExcelFile());
				flipper.initTimeValues();
				flipper.loadMenu();
			}
		}
	}
	
	private void deleteExistingFiles() {
		for (File file : activity.getFilesDir().listFiles()) {
			file.delete();
		}
	}
}
