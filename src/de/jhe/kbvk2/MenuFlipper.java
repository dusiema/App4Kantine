package de.jhe.kbvk2;

import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Locale;
import java.util.TimeZone;

import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.read.biff.BiffException;

import org.apache.http.impl.cookie.DateParseException;
import org.apache.http.impl.cookie.DateUtils;

import android.app.Activity;
import android.app.ProgressDialog;
import android.os.Bundle;
import android.util.Log;
import android.view.GestureDetector;
import android.view.Gravity;
import android.view.Menu;
import android.view.MenuInflater;
import android.view.MenuItem;
import android.view.MotionEvent;
import android.view.View;
import android.view.animation.AnimationUtils;
import android.widget.TextView;
import android.widget.Toast;
import android.widget.ViewFlipper;
import de.jhe.kbvk2.async.DownloadFileAsync;
import de.jhe.kbvk2.data.MenuMetaData;
import de.jhe.kbvk2.gesture.MyGestureDetector;

public class MenuFlipper extends Activity {

	private static final String MENU_URL = "http://a.ndroi.de/android/speiseplan.php";

	private String[] dateFormat = new String[] { "dd.MM.yyyy" };
	private SimpleDateFormat dateFormatter = new SimpleDateFormat(dateFormat[0]);

	private Workbook workbook;
	private HashMap<Integer, MenuMetaData> menuMetaDataMap;
	private static final int MAX_SPALTEN = 8;
	private static final int MAX_ZEILEN = 40;

	private GestureDetector gestureDetector;
	View.OnTouchListener gestureListener;
	private StringBuilder salatNrSB;
	private ViewFlipper flipper;
	private File excelFile = null;
	private Date dateFrom;
	private Date dateUntil;
	private Date dateCurrentlyActive;
//	private Date dateToday = new GregorianCalendar(2011, Calendar.NOVEMBER, 15).getTime();
	private Calendar calendar = Calendar.getInstance(TimeZone.getTimeZone("Europe/Berlin"),
			Locale.GERMANY);;
	private final Date dateToday = calendar.getTime();
	private int dayOfWeekToday;
	private int dayOfWeekFrom;
	private int dayOfWeekUntil;

	private int weekOfYearToday;

	private int weekOfYearFrom;

    private int zeileTagessuppe = 5;
    private int zeileTagessuppeKalorien = 8;
    private int zeileMenu1 = 10;
    private int zeileMenu1Kalorien = 13;
    private int zeileMenu2 = 15;
    private int zeileMenu2Kalorien = 18;
    private int zeileMenu3 = 20;
    private int zeileMenu3Kalorien = 23;
    private int zeileSalat1 = 26;
    private int zeileSalat2 = 27;
    private int zeileDessert1 = 28;
    private int zeileDessert2 = 32;


	/** Called when the activity is first created. */
	@Override
	public void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.menu_flipper);

		// TODO: Remove me again. Just for test purposes. dateToday is set manualy !
		calendar.setTime(dateToday);

		initGestureDetector();
		if (getExcelFile() != null) {
			loadExcel(getExcelFile());
			initTimeValues();
		}
		loadMenu();

	}

	public void initTimeValues() {

		Calendar cal = (Calendar) calendar.clone();
		dayOfWeekToday = cal.get(Calendar.DAY_OF_WEEK);
		weekOfYearToday = cal.get(Calendar.WEEK_OF_YEAR);
		cal.setTime(dateFrom);
		dayOfWeekFrom = cal.get(Calendar.DAY_OF_WEEK);
		weekOfYearFrom = cal.get(Calendar.WEEK_OF_YEAR);
		cal.setTime(dateUntil);
		dayOfWeekUntil = cal.get(Calendar.DAY_OF_WEEK);
		
		if (weekOfYearToday == weekOfYearFrom) {
			if (dayOfWeekToday < dayOfWeekFrom) {
				dateCurrentlyActive = dateFrom;
			} else if (dayOfWeekToday > dayOfWeekUntil) {
				dateCurrentlyActive = dateUntil;
			} else {
				dateCurrentlyActive = dateToday;
			}
		} else {
			dateCurrentlyActive = dateFrom;
		}
		
	}

	private void initGestureDetector() {
		// Gesture detection
		gestureDetector = new GestureDetector(new MyGestureDetector(this));
		gestureListener = new View.OnTouchListener() {
			public boolean onTouch(View v, MotionEvent event) {
				if (gestureDetector.onTouchEvent(event)) {
					return true;
				}
				return false;
			}
		};

	}
	
	/**
	 * L�dt die eigentliche Ansicht und setzt u.a. den aktiven Tag. 
	 */
	public void loadMenu() {

		if (getExcelFile() != null) {
			
			/* 1. 
			 * Set the active child of the flipper to the current date
			 * or to the from date when we are before the first menu date
			 * or to the until date when we are after the last menu date 
			 */
			if (weekOfYearToday == weekOfYearFrom) {
				if ((dayOfWeekToday - 2) >= flipper.getChildCount()) {
					// when we are after friday set to the last child
					flipper.setDisplayedChild(flipper.getChildCount() - 1);
				} else {
					// else set the day
					flipper.setDisplayedChild(dayOfWeekToday - 2);
				}
				
			} else {
				Toast toastMessage = Toast
						.makeText(
								this,
								"Der Speiseplan ist nicht mehr aktuell!",
								Toast.LENGTH_LONG);
				toastMessage.setGravity(Gravity.CENTER, 0, 0);
				toastMessage.show();

			}
			setGauge(flipper.getDisplayedChild());
			setHeaderForDate(dateCurrentlyActive);
			findViewById(R.id.wochentagNext).setClickable(true);
			findViewById(R.id.wochentagPrevious).setClickable(true);
			
		} else {
			Toast toastMessage = Toast
					.makeText(
							this,
							"Kein Speiseplan verf�gbar !\nBitte laden Sie den aktuellen Speiseplan �ber die Funktion \"Lade Speiseplan\".",
							Toast.LENGTH_LONG);
			toastMessage.setGravity(Gravity.CENTER, 0, 0);
			toastMessage.show();
			findViewById(R.id.wochentagNext).setClickable(false);
			findViewById(R.id.wochentagPrevious).setClickable(false);
		}

	}

	public File getExcelFile() {
		if (excelFile == null) {
			File[] files = getFilesDir().listFiles();
			if (files.length > 0) {
				excelFile = files[0];
			}
		}
		return excelFile;
	}
	
	@Override
	public boolean onCreateOptionsMenu(Menu menu) {
		MenuInflater inflater = getMenuInflater();
		inflater.inflate(R.menu.menu, menu);
		return true;
	}

	@Override
	public boolean onOptionsItemSelected(MenuItem item) {
		switch (item.getItemId()) {
		case R.id.loadNewMenu:
			downloadNewMenu();
			break;
		default:
			return super.onOptionsItemSelected(item);
		}
		return true;
	}

	public void downloadNewMenu() {
		ProgressDialog progressDialog = new ProgressDialog(this);
		progressDialog.setProgressStyle(ProgressDialog.STYLE_HORIZONTAL);
		DownloadFileAsync downloadFileAsync = new DownloadFileAsync(this,
				progressDialog);
		downloadFileAsync.execute(MENU_URL);
		progressDialog.setTitle("Download");
		progressDialog.setMessage("Lade Speiseplan...");
		progressDialog.show();
	}

	public void onClickNextDay(View view) {

		flipper.setOutAnimation(AnimationUtils.loadAnimation(this,
				R.anim.anim_translate_left_out));
		flipper.setInAnimation(AnimationUtils.loadAnimation(this,
				R.anim.anim_translate_left_in));
		flipper.showNext();

		flipForward();
	}

	public void onClickPreviousDay(View view) {

		flipper.setOutAnimation(AnimationUtils.loadAnimation(this,
				R.anim.anim_translate_right_out));
		flipper.setInAnimation(AnimationUtils.loadAnimation(this,
				R.anim.anim_translate_right_in));
		flipper.showPrevious();

		flipBackward();
	}

	public void flipForward() {
		setGauge(flipper.getDisplayedChild());

		dateCurrentlyActive = getNextDayDate(dateCurrentlyActive);

		setHeaderForDate(dateCurrentlyActive);
	}

	public void flipBackward() {

		setGauge(flipper.getDisplayedChild());

		dateCurrentlyActive = getPreviousDayDate(dateCurrentlyActive);

		setHeaderForDate(dateCurrentlyActive);
	}

	public boolean isToday(Date date) {
		Calendar calToday = Calendar.getInstance();
		calToday.setTime(dateToday);
		Calendar calCompareTo = Calendar.getInstance();
		calCompareTo.setTime(date);

		if (calToday.get(Calendar.DAY_OF_WEEK) == calCompareTo
				.get(Calendar.DAY_OF_WEEK)) {
			// heute !!!
			return true;
		} else {
			return false;
		}
	}

	public void setGauge(int childIndex) {
		// now change the gauge to the next day
		View gaugeView = flipper.getCurrentView().findViewById(
				R.id.weekdayGauge);
		int gaugeDayViewId = this.getResources().getIdentifier(
				"gauge" + childIndex, "id", this.getPackageName());

		TextView gaugeTextView = (TextView) gaugeView
				.findViewById(gaugeDayViewId);
		gaugeTextView.setBackgroundResource(R.color.magenta);
		gaugeTextView.setTextColor(this.getResources().getColor(R.color.white));

	}

	private void setHeaderForDate(Date date) {
		if (date != null) {
			StringBuilder headerSB = new StringBuilder();
			if (weekOfYearToday == weekOfYearFrom) {
				if ((dayOfWeekToday - 2) == flipper.getDisplayedChild()) {
					if (isToday(date)) {
						headerSB.append("Heute - ");
					}
				}
			}

			TextView wochenttagTextView = (TextView) flipper.getCurrentView()
					.findViewById(R.id.wochentag);
			headerSB.append(dateFormatter.format(dateCurrentlyActive));
			wochenttagTextView.setText(headerSB.toString());
		}
	}

	private Date getNextDayDate(Date date) {
		if (date != null) {
			Calendar calActiveDay = Calendar.getInstance();
			calActiveDay.setTime(date);
			calActiveDay.add(Calendar.DATE, 1);
			date = calActiveDay.getTime();

			Calendar calUntil = Calendar.getInstance();
			calUntil.setTime(dateUntil);

			if (calActiveDay.get(Calendar.DAY_OF_WEEK) > calUntil
					.get(Calendar.DAY_OF_WEEK)) {
				// set to first day again
				date = dateFrom;
			}
		}
		return date;
	}

	private Date getPreviousDayDate(Date date) {
		if (date != null) {
			Calendar calActiveDay = Calendar.getInstance();
			calActiveDay.setTime(date);
			calActiveDay.add(Calendar.DATE, -1);
			date = calActiveDay.getTime();

			Calendar calFrom = Calendar.getInstance();
			calFrom.setTime(dateFrom);

			if (calActiveDay.get(Calendar.DAY_OF_WEEK) < calFrom
					.get(Calendar.DAY_OF_WEEK)) {
				date = dateUntil;
			}
		}
		return date;
	}

	private String getSalatNrText() {
		if (salatNrSB == null) {
			salatNrSB = new StringBuilder();

			String origStringKleinerSalat = menuMetaDataMap.get(zeileSalat1).getMenuNr();
			String origStringGrosserSalat = menuMetaDataMap.get(zeileSalat2).getMenuNr();

            if (origStringKleinerSalat.length() > 0 && origStringGrosserSalat.length() > 0) {
                salatNrSB.append(origStringKleinerSalat.substring(0,
                        origStringKleinerSalat.indexOf(' ')));
                salatNrSB.append(" / ");
                salatNrSB.append(origStringGrosserSalat);
            }
		}

		return salatNrSB.toString();
	}

	public void loadExcel(File file) {
		menuMetaDataMap = new HashMap<Integer, MenuMetaData>();
		flipper = (ViewFlipper) this.findViewById(R.id.menuFlipper);
		
		try {
			WorkbookSettings workbookSettings = new WorkbookSettings();
			workbookSettings.setEncoding("cp1252");
			// workbook =
			// Workbook.getWorkbook(getAssets().open("05_09_2011_bis_11_09_2011_1.xls"),
			// workbookSettings);
			workbook = Workbook.getWorkbook(file, workbookSettings);

			Sheet sheet = workbook.getSheet("Speisenplan");

			for (int spalte = 0; spalte < MAX_SPALTEN; spalte++) {

				// Spalte B
				if (spalte == 1) {
					loadMenuColumn(sheet);
				} else if (spalte == 2) { // Spalte C = Preise
					loadPreiseColumn(sheet);
				} else if (spalte == 3) { // Spalte D = Montag
					loadMondayColumn(sheet);
				} else if (spalte == 4) { // Spalte E = Dienstag
					loadTuesdayColumn(sheet);
				} else if (spalte == 5) { // Spalte F = Mittwoch
					loadWednesdayColumn(sheet);
				} else if (spalte == 6) { // Spalte G = Donnerstag
					loadThursdayColumn(sheet);
				} else if (spalte == 7) { // Spalte H = Freitag
					loadValidity(sheet);
					loadFridayColumn(sheet);
				}
			}

		} catch (BiffException e) {
			Log.d("EXCEL", "Error while loading excel file.", e);
		} catch (IOException e) {
			Log.d("EXCEL", "Error while loading excel file.", e);		}
		workbook.close();
	}

	private void loadValidity(Sheet sheet) {

		String zelle = sheet.getCell(1, 2).getContents();
		String[] dates = zelle.split("-");

		if (dates.length == 2) {

			try {
				dateFrom = DateUtils.parseDate(dates[0].trim(), dateFormat);
				dateUntil = DateUtils.parseDate(dates[1].trim(), dateFormat);
				// dateUntil is sunday - but we need friday, so subtract 2 days
				Calendar cal = Calendar.getInstance();
				cal.setTime(dateUntil);
				cal.add(Calendar.DATE, -2);
				dateUntil = cal.getTime();

			} catch (DateParseException e) {
				Log.d(this.getClass().toString(),
						"Parse exception when reading dates !", e);
			}
		}

		// Set the first day
		TextView wochenttagTextView = (TextView) flipper.getCurrentView()
				.findViewById(R.id.wochentag);

		wochenttagTextView.setText(dateFormatter.format(dateFrom));

	}

	private void loadMenuColumn(Sheet sheet) {
		String zelle;
		MenuMetaData menuMetaData;

		for (int zeile = 0; zeile < MAX_ZEILEN; zeile++) {

			zelle = sheet.getCell(1, zeile).getContents();

			// B7 = Name Tagessuppe
			if (zeile == zeileTagessuppe) {
				menuMetaData = new MenuMetaData();
				menuMetaData.setMenuNr(zelle);
				menuMetaDataMap.put(zeileTagessuppe, menuMetaData);
			}
			// B13 = Name Menu 1
			if (zeile == zeileMenu1) {
				menuMetaData = new MenuMetaData();
				menuMetaData.setMenuNr(zelle);
				menuMetaDataMap.put(zeileMenu1, menuMetaData);
			}
			// B19 = Name Menu 2
			if (zeile == zeileMenu2) {
				menuMetaData = new MenuMetaData();
				menuMetaData.setMenuNr(zelle);
				menuMetaDataMap.put(zeileMenu2, menuMetaData);
			}
			// B25 = Name Menu 3
			if (zeile == zeileMenu3) {
				menuMetaData = new MenuMetaData();
				menuMetaData.setMenuNr(zelle);
				menuMetaDataMap.put(zeileMenu3, menuMetaData);
			}
			// B34 = kleiner Salat
			if (zeile == zeileSalat1) {
				menuMetaData = new MenuMetaData();
				menuMetaData.setMenuNr(zelle);
				menuMetaDataMap.put(zeileSalat1, menuMetaData);
			}
			// B37 = gro�er Salat
			if (zeile == zeileSalat2) {
				menuMetaData = new MenuMetaData();
				menuMetaData.setMenuNr(zelle);
				menuMetaDataMap.put(zeileSalat2, menuMetaData);
			}
			// B40 = Dessert 1
			if (zeile == zeileDessert1) {
				menuMetaData = new MenuMetaData();
				menuMetaData.setMenuNr(zelle);
				menuMetaDataMap.put(zeileDessert1, menuMetaData);
			}
		}
	}

	private void loadPreiseColumn(Sheet sheet) {
		String zelle;

		for (int zeile = 0; zeile < MAX_ZEILEN; zeile++) {
			zelle = sheet.getCell(2, zeile).getContents();

			// C7 = Preis Tagessuppe
			if (zeile == zeileTagessuppe) {
				menuMetaDataMap.get(zeileTagessuppe).setPreis(zelle);
			}
			// C13 = Preis Menu 1
			if (zeile == zeileMenu1) {
				menuMetaDataMap.get(zeileMenu1).setPreis(zelle);
			}
			// C19 = Preis Menu 2
			if (zeile == zeileMenu2) {
				menuMetaDataMap.get(zeileMenu2).setPreis(zelle);
			}
			// C25 = Preis Menu 3
			if (zeile == zeileMenu3) {
				menuMetaDataMap.get(zeileMenu3).setPreis(zelle);
			}
			// C34 = Preis kleiner Salat
			if (zeile == zeileSalat1) {
				menuMetaDataMap.get(zeileSalat1).setPreis(zelle);
			}
			// C37 = Preis gro�er Salat
			if (zeile == zeileSalat2) {
				menuMetaDataMap.get(zeileSalat2).setPreis(zelle);
			}
			// C40 = Preis Desert
			if (zeile == zeileDessert1) {
				menuMetaDataMap.get(zeileDessert1).setPreis(zelle);
			}
		}
	}

	private void loadMondayColumn(Sheet sheet) {
		String zelle;
		View dayView = this.findViewById(R.id.monday);
		View scrollView = dayView.findViewById(R.id.scrollViewWeekDay);
		scrollView.setOnTouchListener(gestureListener);
		((TextView) dayView.findViewById(R.id.wochentag)).setText("Montag");

		for (int zeile = 0; zeile < MAX_ZEILEN; zeile++) {
			zelle = sheet.getCell(3, zeile).getContents();

			// 7 = Tagesuppe - Gericht-Text
			if (zeile == zeileTagessuppe) {
				View view = dayView.findViewById(R.id.gerichtSuppe);
				((TextView) view.findViewById(R.id.gerichtText)).setText(zelle);
				((TextView) view.findViewById(R.id.gerichtNr))
						.setText(menuMetaDataMap.get(zeileTagessuppe).getMenuNr());
				((TextView) view.findViewById(R.id.gerichtPreis))
						.setText(menuMetaDataMap.get(zeileTagessuppe).getPreis());
			}
			// 11 = Tagesuppe - Gericht-Kalorien
			if (zeile == zeileTagessuppeKalorien) {
				View view = dayView.findViewById(R.id.gerichtSuppe);
				((TextView) view.findViewById(R.id.gerichtKcal)).setText(zelle);

			}
			// 13 = Menu1 - Gericht-Text
			if (zeile == zeileMenu1) {
				View view = dayView.findViewById(R.id.gerichtMenu1);
				((TextView) view.findViewById(R.id.gerichtText)).setText(zelle);
				((TextView) view.findViewById(R.id.gerichtNr))
						.setText(menuMetaDataMap.get(zeileMenu1).getMenuNr());
				((TextView) view.findViewById(R.id.gerichtPreis))
						.setText(menuMetaDataMap.get(zeileMenu1).getPreis());

			}
			// 17 = Menu1 - Gericht-Kalorien
			if (zeile == zeileMenu1Kalorien) {
				View view = dayView.findViewById(R.id.gerichtMenu1);
				((TextView) view.findViewById(R.id.gerichtKcal)).setText(zelle);

			}
			// 19 = Menu2 - Gericht-Text
			if (zeile == zeileMenu2) {
				View view = dayView.findViewById(R.id.gerichtMenu2);
				((TextView) view.findViewById(R.id.gerichtText)).setText(zelle);
				((TextView) view.findViewById(R.id.gerichtNr))
						.setText(menuMetaDataMap.get(zeileMenu2).getMenuNr());
				((TextView) view.findViewById(R.id.gerichtPreis))
						.setText(menuMetaDataMap.get(zeileMenu2).getPreis());

			}
			// 23 = Menu2 - Gericht-Kalorien
			if (zeile == zeileMenu2Kalorien) {
				View view = dayView.findViewById(R.id.gerichtMenu2);
				((TextView) view.findViewById(R.id.gerichtKcal)).setText(zelle);

			}
			// 25 = Menu3 - Gericht-Text
			if (zeile == zeileMenu3) {
				View view = dayView.findViewById(R.id.gerichtMenu3);
				((TextView) view.findViewById(R.id.gerichtText)).setText(zelle);
				((TextView) view.findViewById(R.id.gerichtNr))
						.setText(menuMetaDataMap.get(zeileMenu3).getMenuNr());
				((TextView) view.findViewById(R.id.gerichtPreis))
						.setText(menuMetaDataMap.get(zeileMenu3).getPreis());

			}
			// 28 = Menu3 - Gericht-Kalorien
			if (zeile == zeileMenu3Kalorien) {
				View view = dayView.findViewById(R.id.gerichtMenu3);
				((TextView) view.findViewById(R.id.gerichtKcal)).setText(zelle);

			}
			// 34 = kleiner & gro�er Salat
			if (zeile == zeileSalat1) {
				// kleiner Salat
				View view = dayView.findViewById(R.id.gerichtSalat);
				((TextView) view.findViewById(R.id.gerichtText)).setText(zelle);

				((TextView) view.findViewById(R.id.gerichtNr))
						.setText(getSalatNrText());

				String salatPreis = menuMetaDataMap.get(zeileSalat1).getPreis() + " / "
						+ menuMetaDataMap.get(zeileSalat1).getPreis();
				((TextView) view.findViewById(R.id.gerichtPreis))
						.setText(salatPreis);

				((TextView) view.findViewById(R.id.gerichtKcal)).setText("");

			}

			// 40 = Desert
			if (zeile == zeileDessert1) {
				View view = dayView.findViewById(R.id.gerichtDesert);
				((TextView) view.findViewById(R.id.gerichtText)).setText(zelle);
				((TextView) view.findViewById(R.id.gerichtNr))
						.setText(menuMetaDataMap.get(zeileDessert1).getMenuNr());
				((TextView) view.findViewById(R.id.gerichtPreis))
						.setText(menuMetaDataMap.get(zeileDessert1).getPreis());
				((TextView) view.findViewById(R.id.gerichtKcal)).setText("");
			}

		}
	}

	private void loadTuesdayColumn(Sheet sheet) {
		String zelle;
		View dayView = this.findViewById(R.id.tuesday);
		View scrollView = dayView.findViewById(R.id.scrollViewWeekDay);
		scrollView.setOnTouchListener(gestureListener);
		((TextView) dayView.findViewById(R.id.wochentag)).setText("Dienstag");

		for (int zeile = 0; zeile < MAX_ZEILEN; zeile++) {
			zelle = sheet.getCell(4, zeile).getContents();

			// 7 = Tagesuppe - Gericht-Text
			if (zeile == zeileTagessuppe) {
				View view = dayView.findViewById(R.id.gerichtSuppe);
				((TextView) view.findViewById(R.id.gerichtText)).setText(zelle);
				((TextView) view.findViewById(R.id.gerichtNr))
						.setText(menuMetaDataMap.get(zeileTagessuppe).getMenuNr());
				((TextView) view.findViewById(R.id.gerichtPreis))
						.setText(menuMetaDataMap.get(zeileTagessuppe).getPreis());

			}
			// 11 = Tagesuppe - Gericht-Kalorien
			if (zeile == zeileTagessuppeKalorien) {
				View view = dayView.findViewById(R.id.gerichtSuppe);
				((TextView) view.findViewById(R.id.gerichtKcal)).setText(zelle);

			}
			// 13 = Menu1 - Gericht-Text
			if (zeile == zeileMenu1) {
				View view = dayView.findViewById(R.id.gerichtMenu1);
				((TextView) view.findViewById(R.id.gerichtText)).setText(zelle);
				((TextView) view.findViewById(R.id.gerichtNr))
						.setText(menuMetaDataMap.get(zeileMenu1).getMenuNr());
				((TextView) view.findViewById(R.id.gerichtPreis))
						.setText(menuMetaDataMap.get(zeileMenu1).getPreis());

			}
			// 17 = Menu1 - Gericht-Kalorien
			if (zeile == zeileMenu1Kalorien) {
				View view = dayView.findViewById(R.id.gerichtMenu1);
				((TextView) view.findViewById(R.id.gerichtKcal)).setText(zelle);

			}
			// 19 = Menu2 - Gericht-Text
			if (zeile == zeileMenu2) {
				View view = dayView.findViewById(R.id.gerichtMenu2);
				((TextView) view.findViewById(R.id.gerichtText)).setText(zelle);
				((TextView) view.findViewById(R.id.gerichtNr))
						.setText(menuMetaDataMap.get(zeileMenu2).getMenuNr());
				((TextView) view.findViewById(R.id.gerichtPreis))
						.setText(menuMetaDataMap.get(zeileMenu2).getPreis());

			}
			// 23 = Menu2 - Gericht-Kalorien
			if (zeile == zeileMenu2Kalorien) {
				View view = dayView.findViewById(R.id.gerichtMenu2);
				((TextView) view.findViewById(R.id.gerichtKcal)).setText(zelle);

			}
			// 25 = Menu3 - Gericht-Text
			if (zeile == zeileMenu3) {
				View view = dayView.findViewById(R.id.gerichtMenu3);
				((TextView) view.findViewById(R.id.gerichtText)).setText(zelle);
				((TextView) view.findViewById(R.id.gerichtNr))
						.setText(menuMetaDataMap.get(zeileMenu3).getMenuNr());
				((TextView) view.findViewById(R.id.gerichtPreis))
						.setText(menuMetaDataMap.get(zeileMenu3).getPreis());

			}
			// 28 = Menu3 - Gericht-Kalorien
			if (zeile == zeileMenu3Kalorien) {
				View view = dayView.findViewById(R.id.gerichtMenu3);
				((TextView) view.findViewById(R.id.gerichtKcal)).setText(zelle);

			}

			// Folgende Daten stehen nur in der Spalte 3!
			zelle = sheet.getCell(3, zeile).getContents();

			// 34 = kleiner & gro�er Salat
			if (zeile == zeileSalat1) {
				// kleiner Salat
				View view = dayView.findViewById(R.id.gerichtSalat);
				((TextView) view.findViewById(R.id.gerichtText)).setText(zelle);

				((TextView) view.findViewById(R.id.gerichtNr))
						.setText(getSalatNrText());

				String salatPreis = menuMetaDataMap.get(zeileSalat1).getPreis() + " / "
						+ menuMetaDataMap.get(zeileSalat2).getPreis();
				((TextView) view.findViewById(R.id.gerichtPreis))
						.setText(salatPreis);

				((TextView) view.findViewById(R.id.gerichtKcal)).setText("");

			}

			// 40 = Desert
			if (zeile == zeileDessert1) {
				View view = dayView.findViewById(R.id.gerichtDesert);
				((TextView) view.findViewById(R.id.gerichtText)).setText(zelle);
				((TextView) view.findViewById(R.id.gerichtNr))
						.setText(menuMetaDataMap.get(zeileDessert1).getMenuNr());
				((TextView) view.findViewById(R.id.gerichtPreis))
						.setText(menuMetaDataMap.get(zeileDessert1).getPreis());
				((TextView) view.findViewById(R.id.gerichtKcal)).setText("");
			}

		}
	}

	private void loadWednesdayColumn(Sheet sheet) {
		String zelle;
		View dayView = this.findViewById(R.id.wednesday);
		View scrollView = dayView.findViewById(R.id.scrollViewWeekDay);
		scrollView.setOnTouchListener(gestureListener);
		((TextView) dayView.findViewById(R.id.wochentag)).setText("Mittwoch");

		for (int zeile = 0; zeile < MAX_ZEILEN; zeile++) {
			zelle = sheet.getCell(5, zeile).getContents();

			// Zeile 7 = Tagesuppe - Gericht-Text
			if (zeile == zeileTagessuppe) {
				View view = dayView.findViewById(R.id.gerichtSuppe);
				((TextView) view.findViewById(R.id.gerichtText)).setText(zelle);
				((TextView) view.findViewById(R.id.gerichtNr))
						.setText(menuMetaDataMap.get(zeileTagessuppe).getMenuNr());
				((TextView) view.findViewById(R.id.gerichtPreis))
						.setText(menuMetaDataMap.get(zeileTagessuppe).getPreis());

			}
			// Zeile 11 = Tagesuppe - Gericht-Kalorien
			if (zeile == zeileTagessuppeKalorien) {
				View view = dayView.findViewById(R.id.gerichtSuppe);
				((TextView) view.findViewById(R.id.gerichtKcal)).setText(zelle);

			}
			// Zeile 13 = Menu1 - Gericht-Text
			if (zeile == zeileMenu1) {
				View view = dayView.findViewById(R.id.gerichtMenu1);
				((TextView) view.findViewById(R.id.gerichtText)).setText(zelle);
				((TextView) view.findViewById(R.id.gerichtNr))
						.setText(menuMetaDataMap.get(zeileMenu1).getMenuNr());
				((TextView) view.findViewById(R.id.gerichtPreis))
						.setText(menuMetaDataMap.get(zeileMenu1).getPreis());

			}
			// Zeile 17 = Menu1 - Gericht-Kalorien
			if (zeile == zeileMenu1Kalorien) {
				View view = dayView.findViewById(R.id.gerichtMenu1);
				((TextView) view.findViewById(R.id.gerichtKcal)).setText(zelle);

			}
			// Zeile 19 = Menu2 - Gericht-Text
			if (zeile == zeileMenu2) {
				View view = dayView.findViewById(R.id.gerichtMenu2);
				((TextView) view.findViewById(R.id.gerichtText)).setText(zelle);
				((TextView) view.findViewById(R.id.gerichtNr))
						.setText(menuMetaDataMap.get(zeileMenu2).getMenuNr());
				((TextView) view.findViewById(R.id.gerichtPreis))
						.setText(menuMetaDataMap.get(zeileMenu2).getPreis());

			}
			// Zeile 23 = Menu2 - Gericht-Kalorien
			if (zeile == zeileMenu2Kalorien) {
				View view = dayView.findViewById(R.id.gerichtMenu2);
				((TextView) view.findViewById(R.id.gerichtKcal)).setText(zelle);

			}
			// Zeile 25 = Menu3 - Gericht-Text
			if (zeile == zeileMenu3) {
				View view = dayView.findViewById(R.id.gerichtMenu3);
				((TextView) view.findViewById(R.id.gerichtText)).setText(zelle);
				((TextView) view.findViewById(R.id.gerichtNr))
						.setText(menuMetaDataMap.get(zeileMenu3).getMenuNr());
				((TextView) view.findViewById(R.id.gerichtPreis))
						.setText(menuMetaDataMap.get(zeileMenu3).getPreis());

			}
			// Zeile 28 = Menu3 - Gericht-Kalorien
			if (zeile == zeileMenu3Kalorien) {
				View view = dayView.findViewById(R.id.gerichtMenu3);
				((TextView) view.findViewById(R.id.gerichtKcal)).setText(zelle);

			}

			// Folgende Daten stehen nur in der Spalte 3!
			zelle = sheet.getCell(3, zeile).getContents();

			// 34 = kleiner & gro�er Salat
			if (zeile == zeileSalat1) {
				// kleiner Salat
				View view = dayView.findViewById(R.id.gerichtSalat);
				((TextView) view.findViewById(R.id.gerichtText)).setText(zelle);

				((TextView) view.findViewById(R.id.gerichtNr))
						.setText(getSalatNrText());

				String salatPreis = menuMetaDataMap.get(zeileSalat1).getPreis() + " / "
						+ menuMetaDataMap.get(zeileSalat2).getPreis();
				((TextView) view.findViewById(R.id.gerichtPreis))
						.setText(salatPreis);

				((TextView) view.findViewById(R.id.gerichtKcal)).setText("");
			}

			// 40 = Desert
			if (zeile == zeileDessert1) {
				View view = dayView.findViewById(R.id.gerichtDesert);
				((TextView) view.findViewById(R.id.gerichtText)).setText(zelle);
				((TextView) view.findViewById(R.id.gerichtNr))
						.setText(menuMetaDataMap.get(zeileDessert1).getMenuNr());
				((TextView) view.findViewById(R.id.gerichtPreis))
						.setText(menuMetaDataMap.get(zeileDessert1).getPreis());
				((TextView) view.findViewById(R.id.gerichtKcal)).setText("");
			}

		}
	}

	private void loadThursdayColumn(Sheet sheet) {
		String zelle;
		View dayView = this.findViewById(R.id.thursday);
		View scrollView = dayView.findViewById(R.id.scrollViewWeekDay);
		scrollView.setOnTouchListener(gestureListener);
		((TextView) dayView.findViewById(R.id.wochentag)).setText("Donnerstag");

		for (int zeile = 0; zeile < MAX_ZEILEN; zeile++) {
			zelle = sheet.getCell(6, zeile).getContents();

			// D7 = Tagesuppe - Gericht-Text
			if (zeile == zeileTagessuppe) {
				View view = dayView.findViewById(R.id.gerichtSuppe);
				((TextView) view.findViewById(R.id.gerichtText)).setText(zelle);
				((TextView) view.findViewById(R.id.gerichtNr))
						.setText(menuMetaDataMap.get(zeileTagessuppe).getMenuNr());
				((TextView) view.findViewById(R.id.gerichtPreis))
						.setText(menuMetaDataMap.get(zeileTagessuppe).getPreis());

			}
			// D11 = Tagesuppe - Gericht-Kalorien
			if (zeile == zeileTagessuppeKalorien) {
				View view = dayView.findViewById(R.id.gerichtSuppe);
				((TextView) view.findViewById(R.id.gerichtKcal)).setText(zelle);

			}
			// D13 = Menu1 - Gericht-Text
			if (zeile == zeileMenu1) {
				View view = dayView.findViewById(R.id.gerichtMenu1);
				((TextView) view.findViewById(R.id.gerichtText)).setText(zelle);
				((TextView) view.findViewById(R.id.gerichtNr))
						.setText(menuMetaDataMap.get(zeileMenu1).getMenuNr());
				((TextView) view.findViewById(R.id.gerichtPreis))
						.setText(menuMetaDataMap.get(zeileMenu1).getPreis());

			}
			// D17 = Menu1 - Gericht-Kalorien
			if (zeile == zeileMenu1Kalorien) {
				View view = dayView.findViewById(R.id.gerichtMenu1);
				((TextView) view.findViewById(R.id.gerichtKcal)).setText(zelle);

			}
			// D19 = Menu2 - Gericht-Text
			if (zeile == zeileMenu2) {
				View view = dayView.findViewById(R.id.gerichtMenu2);
				((TextView) view.findViewById(R.id.gerichtText)).setText(zelle);
				((TextView) view.findViewById(R.id.gerichtNr))
						.setText(menuMetaDataMap.get(zeileMenu2).getMenuNr());
				((TextView) view.findViewById(R.id.gerichtPreis))
						.setText(menuMetaDataMap.get(zeileMenu2).getPreis());

			}
			// D23 = Menu2 - Gericht-Kalorien
			if (zeile == zeileMenu2Kalorien) {
				View view = dayView.findViewById(R.id.gerichtMenu2);
				((TextView) view.findViewById(R.id.gerichtKcal)).setText(zelle);

			}
			// D25 = Menu3 - Gericht-Text
			if (zeile == zeileMenu3) {
				View view = dayView.findViewById(R.id.gerichtMenu3);
				((TextView) view.findViewById(R.id.gerichtText)).setText(zelle);
				((TextView) view.findViewById(R.id.gerichtNr))
						.setText(menuMetaDataMap.get(zeileMenu3).getMenuNr());
				((TextView) view.findViewById(R.id.gerichtPreis))
						.setText(menuMetaDataMap.get(zeileMenu3).getPreis());

			}
			// D28 = Menu3 - Gericht-Kalorien
			if (zeile == zeileMenu3Kalorien) {
				View view = dayView.findViewById(R.id.gerichtMenu3);
				((TextView) view.findViewById(R.id.gerichtKcal)).setText(zelle);

			}

			// Folgende Daten stehen nur in der Spalte 3!
			zelle = sheet.getCell(3, zeile).getContents();

			// 34 = kleiner & gro�er Salat
			if (zeile == zeileSalat1) {
				// kleiner Salat
				View view = dayView.findViewById(R.id.gerichtSalat);

				((TextView) view.findViewById(R.id.gerichtText)).setText(zelle);
				((TextView) view.findViewById(R.id.gerichtNr))
						.setText(getSalatNrText());

				String salatPreis = menuMetaDataMap.get(zeileSalat1).getPreis() + " / "
						+ menuMetaDataMap.get(zeileSalat2).getPreis();
				((TextView) view.findViewById(R.id.gerichtPreis))
						.setText(salatPreis);

				((TextView) view.findViewById(R.id.gerichtKcal)).setText("");
			}

			// 40 = Desert
			if (zeile == zeileDessert1) {
				View view = dayView.findViewById(R.id.gerichtDesert);
				((TextView) view.findViewById(R.id.gerichtText)).setText(zelle);
				((TextView) view.findViewById(R.id.gerichtNr))
						.setText(menuMetaDataMap.get(zeileDessert1).getMenuNr());
				((TextView) view.findViewById(R.id.gerichtPreis))
						.setText(menuMetaDataMap.get(zeileDessert1).getPreis());
				((TextView) view.findViewById(R.id.gerichtKcal)).setText("");
			}

		}
	}

	private void loadFridayColumn(Sheet sheet) {
		String zelle;
		View dayView = this.findViewById(R.id.friday);
		View scrollView = dayView.findViewById(R.id.scrollViewWeekDay);
		scrollView.setOnTouchListener(gestureListener);
		((TextView) dayView.findViewById(R.id.wochentag)).setText("Freitag");

		for (int zeile = 0; zeile < MAX_ZEILEN; zeile++) {
			zelle = sheet.getCell(7, zeile).getContents();

			// D7 = Tagesuppe - Gericht-Text
			if (zeile == zeileTagessuppe) {
				View view = dayView.findViewById(R.id.gerichtSuppe);
				((TextView) view.findViewById(R.id.gerichtText)).setText(zelle);
				((TextView) view.findViewById(R.id.gerichtNr))
						.setText(menuMetaDataMap.get(zeileTagessuppe).getMenuNr());
				((TextView) view.findViewById(R.id.gerichtPreis))
						.setText(menuMetaDataMap.get(zeileTagessuppe).getPreis());

			}
			// D11 = Tagesuppe - Gericht-Kalorien
			if (zeile == zeileTagessuppeKalorien) {
				View view = dayView.findViewById(R.id.gerichtSuppe);
				((TextView) view.findViewById(R.id.gerichtKcal)).setText(zelle);

			}
			// D13 = Menu1 - Gericht-Text
			if (zeile == zeileMenu1) {
				View view = dayView.findViewById(R.id.gerichtMenu1);
				((TextView) view.findViewById(R.id.gerichtText)).setText(zelle);
				((TextView) view.findViewById(R.id.gerichtNr))
						.setText(menuMetaDataMap.get(zeileMenu1).getMenuNr());
				((TextView) view.findViewById(R.id.gerichtPreis))
						.setText(menuMetaDataMap.get(zeileMenu1).getPreis());

			}
			// D17 = Menu1 - Gericht-Kalorien
			if (zeile == zeileMenu1Kalorien) {
				View view = dayView.findViewById(R.id.gerichtMenu1);
				((TextView) view.findViewById(R.id.gerichtKcal)).setText(zelle);

			}
			// D19 = Menu2 - Gericht-Text
			if (zeile == zeileMenu2) {
				View view = dayView.findViewById(R.id.gerichtMenu2);
				((TextView) view.findViewById(R.id.gerichtText)).setText(zelle);
				((TextView) view.findViewById(R.id.gerichtNr))
						.setText(menuMetaDataMap.get(zeileMenu2).getMenuNr());
				((TextView) view.findViewById(R.id.gerichtPreis))
						.setText(menuMetaDataMap.get(zeileMenu2).getPreis());

			}
			// D23 = Menu2 - Gericht-Kalorien
			if (zeile == zeileMenu2Kalorien) {
				View view = dayView.findViewById(R.id.gerichtMenu2);
				((TextView) view.findViewById(R.id.gerichtKcal)).setText(zelle);

			}
			// D25 = Menu3 - Gericht-Text
			if (zeile == zeileMenu3) {
				View view = dayView.findViewById(R.id.gerichtMenu3);
				((TextView) view.findViewById(R.id.gerichtText)).setText(zelle);
				((TextView) view.findViewById(R.id.gerichtNr))
						.setText(menuMetaDataMap.get(zeileMenu3).getMenuNr());
				((TextView) view.findViewById(R.id.gerichtPreis))
						.setText(menuMetaDataMap.get(zeileMenu3).getPreis());

			}
			// D28 = Menu3 - Gericht-Kalorien
			if (zeile == zeileMenu3Kalorien) {
				View view = dayView.findViewById(R.id.gerichtMenu3);
				((TextView) view.findViewById(R.id.gerichtKcal)).setText(zelle);

			}

			// Folgende Daten stehen nur in der Spalte 3!
			zelle = sheet.getCell(3, zeile).getContents();

			// 34 = kleiner & gro�er Salat
			if (zeile == zeileSalat1) {
				// kleiner Salat
				View view = dayView.findViewById(R.id.gerichtSalat);
				((TextView) view.findViewById(R.id.gerichtText)).setText(zelle);

				((TextView) view.findViewById(R.id.gerichtNr))
						.setText(getSalatNrText());

				String salatPreis = menuMetaDataMap.get(zeileSalat1).getPreis() + " / "
						+ menuMetaDataMap.get(zeileSalat2).getPreis();
				((TextView) view.findViewById(R.id.gerichtPreis))
						.setText(salatPreis);

				((TextView) view.findViewById(R.id.gerichtKcal)).setText("");
			}

			// 40 = Desert
			if (zeile == zeileDessert1) {
				View view = dayView.findViewById(R.id.gerichtDesert);
				((TextView) view.findViewById(R.id.gerichtText)).setText(zelle);
				((TextView) view.findViewById(R.id.gerichtNr))
						.setText(menuMetaDataMap.get(zeileDessert1).getMenuNr());
				((TextView) view.findViewById(R.id.gerichtPreis))
						.setText(menuMetaDataMap.get(zeileDessert1).getPreis());
				((TextView) view.findViewById(R.id.gerichtKcal)).setText("");
			}

		}
	}

}