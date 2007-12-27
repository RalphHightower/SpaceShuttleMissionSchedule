/*
    NasaTvScheduleImport.  This program reads the NASA TV Schedule in Excel
    Format for the Space Shuttle and transfers the entries into Microsoft
    Outlook Calendar as Appointment items.

    Copyright (C) 2007  Ralph M. Hightower, Jr (Permanent Vacations)

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.

    mailto:Ralph.Hightower@gmail.com
 * Ralph Hightower
 * Chapin, SC 29036
*/
/*
 * Revision History
 * 20071227 - Ralph Hightower
 *		Fixed nagging problems with pairing crew sleep with crew wakeup calls; 
 *		Tightened up rules for Happy New Year Routine for Gregorian calendar.
 * 20071226 - Ralph Hightower
 *		STS-118 schedule has entries with a site of " " causing a null site to be entered in the schedule
 */
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Security.Permissions;
using System.Windows.Forms;
using InteropExcel = Microsoft.Office.Interop.Excel;
using ToolsExcel = Microsoft.Office.Tools.Excel;
using Microsoft.Msdn.BclTeam;

[assembly: CLSCompliant(true)]
namespace PermanentVacations.Nasa.Sts.Schedule
{
	/// <summary>
	/// Enum to interpret the different types of rows in the Space Shuttle TV Schedule spreadsheet
	/// empty: blank
	/// columnHeading: for the row containing the column header for the schedule events
	/// dateHeading: changes the Date
	/// flightDayHeading: header for the Flight Dat
	/// scheduleEntry: the event with start and end times
	///		The Subject may be on multiple lines in the same column, other column entries will be
	///		blank if the subject is continued
	/// definitionOfTerms: end of file, definitions are skipped
	/// </summary>
	public enum ScheduleType
	{
		empty, columnHeading, dateHeading, flightDayHeading, scheduleEntry,
			definitionOfTerms, error
	};

	#region NasaStsTVSchedule
	/// <summary>
	/// Interprets the Excel cells in the NASA STS TV Schedule
	/// </summary>
	public class NasaStsTVSchedule : Object
	{
		/// <summary>
		/// Carriage Return, Line Feed
		/// </summary>
		private const string CRLF = "\r\n";

		#region Frequently Used Regular Expressions - Compiled
		//	Switching Regular Expressions to Regex.Compiled brought the average processing of an
		//	Excel spreadsheet from 45 seconds to 15 seconds which is a significant improvement in speed

		/// <summary>
		/// Compiled Regular Expression for Date Header entry
		/// Date Header is DayOfWeek, Month missionDay in uppercase
		/// </summary>
		private Regex crgDateHeader = null;
		/// <summary>
		/// Compiled Regular Expression for Date Header entry
		/// Getter creates compiled instance of Regex if it hasn't been created yet.
		/// Setter does not check for nulls since it will be used in ~NasaStsTVSchedule
		/// </summary>
		private Regex rgDateHeader
		{
			get
			{
				if (crgDateHeader == null)
				{
					crgDateHeader = new Regex(Properties.Resources.RGX_DATE_HEADER, RegexOptions.Compiled |
						RegexOptions.CultureInvariant);
				}
				return (crgDateHeader);
			}
			set
			{
				if (value == null)
					crgDateHeader = null;
				else
					crgDateHeader = value;
			}
		}
		/// <summary>
		/// Compiled Regular Expression for EVA Activity
		/// Eva Activity is EVA #d BEGINS or ENDs
		/// </summary>
		private Regex crgEvaActivity = null;
		/// <summary>
		/// Compiled Regular Expression for EVA Activity
		/// Getter creates compiled instance of Regex if it hasn't been created yet.
		/// Setter does not check for nulls since it will be used in ~NasaStsTVSchedule
		/// </summary>
		private Regex rgEvaActivity
		{
			get
			{
				if (crgEvaActivity == null)
				{
					crgEvaActivity = new Regex(Properties.Resources.RGX_EVA_ACTIVITY, RegexOptions.Compiled |
						RegexOptions.CultureInvariant);
				}
				return (crgEvaActivity);
			}
			set
			{
				if (value == null)
					crgEvaActivity = null;
				else
					crgEvaActivity = null;
			}
		}
		/// <summary>
		/// Compiled Regular Expression for Flight Day Header
		/// Flight Day Header is: FD #d
		/// </summary>
		private Regex crgFlightDayHeader = null;
		/// <summary>
		/// Compiled Regular Expression for Flight Day Header
		/// Getter creates compiled instance of Regex if it hasn't been created yet.
		/// Setter does not check for nulls since it will be used in ~NasaStsTVSchedule
		/// </summary>
		private Regex rgFlightDayHeader
		{
			get
			{
				if (crgFlightDayHeader == null)
				{
					crgFlightDayHeader = new Regex(Properties.Resources.RGX_FLIGHDAY_HEADER, RegexOptions.Compiled |
						RegexOptions.CultureInvariant);
				}
				return (crgFlightDayHeader);
			}
			set
			{
				if (value == null)
					crgFlightDayHeader = null;
				else
					crgFlightDayHeader = value;
			}
		}
		/// <summary>
		/// Compiled Regular Expression for the Flight Day Highlights schedule entry
		/// Format is: FLIGHT DAY #d HIGHLIGHTS
		/// </summary>
		private Regex crgFlightDayHighlights;
		/// Compiled Regular Expression for Flight Day Highlights
		/// Getter creates compiled instance of Regex if it hasn't been created yet.
		/// Setter does not check for nulls since it will be used in ~NasaStsTVSchedule
		private Regex rgFlightDayHighlights
		{
			get
			{
				if (crgFlightDayHighlights == null)
				{
					crgFlightDayHighlights = new Regex(Properties.Resources.TM_RG_FLIGHT_DAY_HIGHLIGHTS,
						RegexOptions.Compiled | RegexOptions.CultureInvariant);
				}
				return (crgFlightDayHighlights);
			}
			set
			{
				if (value == null)
					crgFlightDayHighlights = null;
				else
					crgFlightDayHighlights = value;
			}
		}
		/// <summary>
		/// Compiled Regular Expression for ISS Crew Sleep Activity
		/// ISS Crew Sleep Activity is: Optional: "Shuttle name / ", Required: ISS CREW SLEEP BEGINS or ENDS
		/// </summary>
		private Regex crgIssCrewSleepActivity = null;
		/// <summary>
		/// Compiled Regular Expression for ISS Crew Sleep Activity
		/// Getter creates compiled instance of Regex if it hasn't been created yet.
		/// Setter does not check for nulls since it will be used in ~NasaStsTVSchedule
		/// </summary>
		private Regex rgIssCrewSleepActivity
		{
			get
			{
				if (crgIssCrewSleepActivity == null)
				{
					crgIssCrewSleepActivity = new Regex(Properties.Resources.RGX_ISS_CREW_SLEEP_ACTIVITY,
						RegexOptions.Compiled | RegexOptions.CultureInvariant);
				}
				return (crgIssCrewSleepActivity);
			}
			set
			{
				if (value == null)
					crgIssCrewSleepActivity = null;
				else
					crgIssCrewSleepActivity = value;
			}
		}
		/// <summary>
		/// Regular Express for Shuttle Crew Sleep Activity
		/// Shuttle Crew Sleep Activity is Required: Shuttle name, Optional: / ISS, Required: CREW SLEEP BEGINS or ENDS
		/// </summary>
		private Regex crgShuttleCrewSleepActivity = null;
		/// <summary>
		/// Regular Express for Shuttle Crew Sleep Activity
		/// Getter creates compiled instance of Regex if it hasn't been created yet.
		/// Setter does not check for nulls since it will be used in ~NasaStsTVSchedule
		/// </summary>
		private Regex rgShuttleCrewSleepActivity
		{
			get
			{
				if (crgShuttleCrewSleepActivity == null)
				{
					crgShuttleCrewSleepActivity = new Regex(Properties.Resources.RGX_SHUTTLE_CREW_SLEEP_ACTIVITY,
						RegexOptions.Compiled | RegexOptions.CultureInvariant);
				}
				return (crgShuttleCrewSleepActivity);
			}
			set
			{
				if (value == null)
					crgShuttleCrewSleepActivity = null;
				else
					crgShuttleCrewSleepActivity = value;
			}
		}
		#endregion	//	Frequently Used Regular Expressions - Compiled

		#region PermanentVacations.Nasa.Sts.Schedule Variables with Getters/Setters
		/// <summary>
		/// Captures exceptions to pass back to application
		///	Invalid File Formats are defined as:
		/// 1)	No creation or revision date in the first few rows before the Date Header
		///		to get the year of the mission
		/// 2)	Print_Area is not defined as a name in the spreadsheet containing the range
		///		of cells containing the schedule
		/// </summary>
		InvalidFileFormatException processingException = null;
		/// <summary>
		/// Hold state of exceptions encountered during processing
		/// </summary>
		ScheduleType errorProcessing = ScheduleType.empty;
		/// <summary>
		/// Returns true if an error has been encountered
		/// </summary>
		private bool ErrorProcessed
		{
			get { return (errorProcessing == ScheduleType.error); }
		}
		/// <summary>
		/// Getter/Setter for InvalidFileFormatException
		/// </summary>
		private InvalidFileFormatException ProcessingError
		{
			get { return (processingException); }
			set
			{
				if (value == null)
					processingException = null;
				else
				{
					processingException = value;
					errorProcessing = ScheduleType.error;
				}
			}
		}
		/// <summary>
		/// Year of schedule
		/// </summary>
		private int missionYear;
		/// <summary>
		/// Getters/Setters for missionYear
		/// </summary>
		private int Year
		{
			get { return (missionYear); }
			set { missionYear = value; }
		}
		/// <summary>
		/// Month of mission
		/// </summary>
		private int missionMonth;
		/// <summary>
		/// Getters/Setters for missionMonth (handles the New Year)
		/// </summary>
		private int Month
		{
			get { return (missionMonth); }
			set
			{
				//	Revision Date has not been found before a Date Header was processed
				//	Should throw an exception that a revision date has not been encountered before the Date Header
				if (Year == 0)
				{
					string explanation = Properties.Resources.INVALIDFILE_NO_REVISION_DATE;
					throw new InvalidFileFormatException(String.Format(explanation, NasaTVScheduleFile));
				}
				if (missionMonth > 0)
				{
					//	If missionMonth has already been set and the new missionMonth is less than the current missionMonth ...
					//	HAPPY NEW YEAR!
					if (value < missionMonth)
					{
						if ((missionMonth == 12) && (value == 1))
						{
							Year++;
						}
					}
				}
				missionMonth = value;
			}
		}
		/// <summary>
		/// Day of the mission
		/// </summary>
		private int missionDay;
		/// <summary>
		/// Getters/Setters for the missionDay of the mission
		/// </summary>
		private int Day
		{
			get { return (missionDay); }
			set { missionDay = value; }
		}

		/// <summary>
		/// Enum for the type of Excel interface used (Microsoft.Office.Interop.Excel or Microsoft.Office.Tools.Excel)
		/// Microsoft.Office.Interop.Excel was the first to get working to open and read files
		/// Still working on figuring out how to do the same with Microsoft.Office.Tools.Excel
		/// </summary>
		public enum ExcelInterface { InteropExcel, ToolsExcel };
		/// <summary>
		/// State variable for the type of Excel interface
		/// </summary>
		private ExcelInterface m_excelInterface;
		/// <summary>
		/// Gets/Sets the type of Excel interface
		/// </summary>
		private ExcelInterface ExcelTypeInterface
		{
			get { return (m_excelInterface); }
			set { m_excelInterface = value; }
		}

		/// <summary>
		/// Class using Microsoft.Office.Interop.Excel
		/// </summary>
		InteropExcelInterface interopExcel;
		/// <summary>
		/// Getter/Setter fir Microsoft.Office.Interop.Excel Class
		/// </summary>
		private InteropExcelInterface InteropExcelIF
		{
			get
			{
				if (interopExcel == null)
				{
					interopExcel = new InteropExcelInterface();
				}
				return (interopExcel);
			}
			set
			{
				if (value == null)
				{
					interopExcel.Dispose();
					interopExcel = null;
				}

				else
					interopExcel = value;
			}
		}

		/// <summary>
		/// Class using Microsoft.Office.Tools.Excel
		/// </summary>
		private ToolsExcelInterface toolsExcelInterface;
		/// <summary>
		/// Getter/Setter for Microsoft.Office.Interop.Excel Class
		/// </summary>
		private ToolsExcelInterface ToolsExcelIF
		{
			get
			{
				if (toolsExcelInterface == null)
				{
					toolsExcelInterface = new ToolsExcelInterface();
				}
				return (toolsExcelInterface);
			}
			set
			{
				if (value == null)
					toolsExcelInterface = null;
				else
					toolsExcelInterface = value;
			}
		}

		/// <summary>
		/// Array to hold Excel Range named "Print Area"
		/// </summary>
		private System.Array m_sysarrTvScheduleCells;
		/// <summary>
		/// Gets/Sets Array holding Excel Range "Print Area"
		/// </summary>
		private System.Array TvScheduleCells
		{
			get { return (m_sysarrTvScheduleCells); }
			set
			{
				if (value != null)
					m_sysarrTvScheduleCells = value;
				else
					m_sysarrTvScheduleCells = null;
			}
		}

		/// <summary>
		/// Returns percent of spreadsheet processed
		/// </summary>
		public Double PercentComplete
		{
			get
			{
				Double percent = 0.0;
				if (RowCount > 0)
					percent = CurrentRow / RowCount;
				return (percent);
			}
		}

		/// <summary>
		/// Returns the Current Row
		/// </summary>
		public int CurrentEntry
		{
			get { return (CurrentRow); }
		}

		/// <summary>
		/// Returns the number of rows
		/// Getter exposed to public for progress tracking
		/// </summary>
		public int NumberRows
		{
			get { return (RowCount); }
		}

		/// <summary>
		/// Number of Rows in spreadsheet
		/// </summary>
		private int m_RowCount;
		/// <summary>
		/// Gets/Sets Number of Rows in spreadsheet
		/// Private use only
		/// </summary>
		private int RowCount
		{
			get { return (m_RowCount); }
			set { m_RowCount = value; }
		}

		/// <summary>
		/// Number of Columns in Spreadsheet
		/// </summary>
		private int m_ColumnCount;
		/// <summary>
		/// Gets/Sets Number of Columns in Spreadsheet
		/// </summary>
		private int ColumnCount
		{
			get { return (m_ColumnCount); }
			set { m_ColumnCount = value; }
		}

		/// <summary>
		/// Name of Nasa Sts TV Schedule file
		/// </summary>
		private string m_NasaTVScheduleFile;
		/// <summary>
		/// Gets/Sets Name of Nasa Sts TV Schedule file
		/// </summary>
		private string NasaTVScheduleFile
		{
			get { return (m_NasaTVScheduleFile); }
			set
			{
				if (value != null)
					m_NasaTVScheduleFile = value;
				else
					m_NasaTVScheduleFile = null;
			}
		}

		/// <summary>
		/// Returns true if Excel has successfully opened the NASA STS TV Schedule file
		/// </summary>
		public bool SuccessfullyOpened
		{
			get
			{
				bool opened = false;
				switch (ExcelTypeInterface)
				{
					case ExcelInterface.InteropExcel:
						if (InteropExcelIF != null)
							opened = InteropExcelIF.SuccessfullyOpened;
						break;
					case ExcelInterface.ToolsExcel:
						if (ToolsExcelIF != null)
							opened = ToolsExcelIF.SuccessfullyOpened;
						break;
				}
				return (opened);
			}
		}

		/// <summary>
		/// Internal array of Weekdays (Uppercase)
		/// </summary>
		private string[] Days = {Properties.Resources.NASA_DOW_SUNDAY, Properties.Resources.NASA_DOW_MONDAY,
			Properties.Resources.NASA_DOW_TUESDAY, Properties.Resources.NASA_DOW_WEDNESDAY,
			Properties.Resources.NASA_DOW_THURSDAY, Properties.Resources.NASA_DOW_FRIDAY,
			Properties.Resources.NASA_DOW_SATURDAY};
		/// <summary>
		/// Internal array of Months (Uppercase)
		/// </summary>
		private string[] Months = {Properties.Resources.NASA_MON_JANUARY, Properties.Resources.NASA_MON_FEBRUARY,
			Properties.Resources.NASA_MON_MARCH, Properties.Resources.NASA_MON_APRIL,
			Properties.Resources.NASA_MON_MAY, Properties.Resources.NASA_MON_JUNE,
			Properties.Resources.NASA_MON_JULY, Properties.Resources.NASA_MON_AUGUST,
			Properties.Resources.NASA_MON_SEPTEMBER, Properties.Resources.NASA_MON_OCTOBER,
			Properties.Resources.NASA_MON_NOVEMBER, Properties.Resources.NASA_MON_DECEMBER};

		//  TO DO:
		//      Develop a method to get the comments from the Resources file (Ralph Hightower)

		/// <summary>
		/// Some Nasa events do not last to the next event in the schedule
		/// This is just a guesstimate of how long some events last.
		/// </summary>
		//  TM_RG means Time Event, Regular Expression
		private string[,] EventTimes = { { Properties.Resources.TM_RG_FLIGHT_DAY_HIGHLIGHTS, "00:45:00" },
			{ Properties.Resources.TM_VIDEO_FILE, "01:00:00" },
			{ Properties.Resources.TM_ISS_FLIGHT_DIRECTOR_UPDATE, "00:15:00" },
			{ Properties.Resources.TM_MISSION_STATUS_BRIEFING, "00:45:00" },
			{ Properties.Resources.TM_PAO_EVENT, "00:15:00" },
			{ Properties.Resources.TM_NEWS_CONFERENCE, "01:00:00" },
			{ Properties.Resources.TM_POST_MMT_BRIEFING, "00:45:00" },
			// ReadAhead will detect wake up period
			//{ Properties.Resources.TM_CREW_SLEEP_BEGINS, "08:30:00" },
			{ Properties.Resources.TM_INTERVIEW, "00:15:00" },
			{ Properties.Resources.TM_CHANGE_COMMAND_CEREMONY, "00:20:00"},
			{ Properties.Resources.TM_COUNTDOWN_STATUS_BRIEFING, "01:00:00"},
			{ Properties.Resources.TM_BRIEFING, "00:30:00" },
			{ Properties.Resources.TM_DEORBIT_BURN, "00:02:30" }
		};

		/// <summary>
		/// Internal array holding list of Windows Time Zones from the registry
		/// </summary>
		private TimeZoneInfo[] m_TimeZoneInfo;
		/// <summary>
		/// Gets/Sets the internal array of Windows Time Zones
		/// </summary>
		private TimeZoneInfo[] TZInfo
		{
			get { return (m_TimeZoneInfo); }
			set
			{
				if (value != null)
					m_TimeZoneInfo = value;
				else
					m_TimeZoneInfo = null;
			}
		}
		/// <summary>
		/// The Current Row of the schedule being examined
		/// </summary>
		private int m_Current_Row;
		/// <summary>
		/// Gets/Sets the Current Row count
		/// Checks to make sure the Current Row does not go out of bounds
		/// </summary>
		private int CurrentRow
		{
			get { return (m_Current_Row); }
			set
			{
				m_Current_Row = value;
				if ((RowCount > 0) && (CurrentRow > RowCount))
				{
					IsEOF = true;
				}
			}
		}
		/// <summary>
		/// Column of the Orbit Header
		/// </summary>
		private int m_colHdrOrbit;
		/// <summary>
		/// Get/Sets the Orbit Header Column
		/// </summary>
		private int OrbitColumnHeader
		{
			get { return (m_colHdrOrbit); }
			set { m_colHdrOrbit = value; }
		}
		/// <summary>
		/// Indicates if Nasa changed the particular entry in this revision
		/// </summary>
		private bool m_Changed;
		/// <summary>
		/// Gets the change indicator for the event
		/// </summary>
		/// <returns>true if Nasa revised event in this revision</returns>
		public bool Revised()
		{
			return (m_Changed);
		}
		/// <summary>
		/// Gets/Sets the Nasa event Changed flag in this revision
		/// </summary>
		private bool Changed
		{
			get { return (m_Changed); }
			set { m_Changed = value; }
		}
		/// <summary>
		/// Column of the Subject
		/// </summary>
		private int m_colHdrSubject;
		/// <summary>
		/// Gets/Sets the Column of the Subject
		/// </summary>
		private int SubjectColumnHeader
		{
			get { return (m_colHdrSubject); }
			set { m_colHdrSubject = value; }
		}
		/// <summary>
		/// Column for the Site of the event
		/// </summary>
		private int m_colHdrSite;
		/// <summary>
		/// Gets/Sets the Column for the Site
		/// </summary>
		private int SiteColumHeader
		{
			get { return (m_colHdrSite); }
			set { m_colHdrSite = value; }
		}
		/// <summary>
		/// Column for the Flight Day
		/// </summary>
		private int m_colHdrFD;
		/// <summary>
		/// Gets/Sets the column for the Flight Day
		/// </summary>
		private int FlightDayColumnHeader
		{
			get { return (m_colHdrFD); }
			set { m_colHdrFD = value; }
		}

		/// <summary>
		/// Mission Elapsed Time of mission
		///		int day
		///		TimeSpan time
		/// </summary>
		private MissionDuration m_MissionDurationTime;
		/// <summary>
		/// Getter/Setter for Mission Elapsed Time
		/// </summary>
		private MissionDuration MissionDurationTime
		{
			get
			{
				if (m_MissionDurationTime == null)
					m_MissionDurationTime = new MissionDuration();
				return (m_MissionDurationTime);
			}
			set
			{
				if (value == null)
					m_MissionDurationTime = null;
				else
					m_MissionDurationTime = value;
			}
		}

		/// <summary>
		/// Column for the Mission Elapsed Time
		/// </summary>
		private int m_colHdrMissionElaspedTime;
		/// <summary>
		/// Gets/Sets the column for the Mission Elapsed Time
		/// </summary>
		private int MissionElapsedTimeColumnHeader
		{
			get { return (m_colHdrMissionElaspedTime); }
			set { m_colHdrMissionElaspedTime = value; }
		}
		/// <summary>
		/// Column for Central Time Zone
		/// </summary>
		private int m_colHdrCentralTime;
		/// <summary>
		/// Gets/Sets the column for the Central Time Zone
		/// </summary>
		private int CentralTimeColumnHeader
		{
			get { return (m_colHdrCentralTime); }
			set { m_colHdrCentralTime = value; }
		}
		/// <summary>
		/// Column for Eastern Time Zone
		/// </summary>
		private int m_colHdrEasternTime;
		/// <summary>
		/// Gets/Sets column for Eastern Time Zone
		/// </summary>
		private int EasternTimeColumnHeader
		{
			get { return (m_colHdrEasternTime); }
			set { m_colHdrEasternTime = value; }
		}
		/// <summary>
		/// Column for GMT
		/// </summary>
		private int m_colHdrGreenwichMeanTime;
		/// <summary>
		/// Gets/Sets column for GMT
		/// </summary>
		private int GreenwichMeanTimeColumnHeader
		{
			get { return (m_colHdrGreenwichMeanTime); }
			set { m_colHdrGreenwichMeanTime = value; }
		}
		/// <summary>
		/// Is Current Event in Daylight Savings Time
		/// </summary>
		private bool m_isDaylightSavingsTime;
		/// <summary>
		/// Gets/Sets Daylight Savings Time status
		/// </summary>
		private bool DaylightSavingsTime
		{
			get { return (m_isDaylightSavingsTime); }
			set { m_isDaylightSavingsTime = value; }
		}
		/// <summary>
		/// Has End of Spreadsheet been reached?
		/// </summary>
		private bool m_isEOF;
		/// <summary>
		/// Returns status of End of Spreadsheet
		/// </summary>
		/// <returns>true if the end of the spreadsheet has been reached</returns>
		public bool EOF()
		{
			return (IsEOF);
		}
		/// <summary>
		/// Gets/Sets End of Spreadsheet status
		/// </summary>
		private bool IsEOF
		{
			get { return (m_isEOF); }
			set { m_isEOF = value; }
		}
		/// <summary>
		/// Indicates if Shuttle is docked to ISS
		/// Not reliable for revisions after shuttle has docked
		/// </summary>
		private bool m_isDocked;
		/// <summary>
		/// Gets status of shuttle docking to ISS (not reliable after docking with revised schedules)
		/// </summary>
		/// <returns>true if the shuttle is docked to ISS (not reliable after docking with revised schedules)</returns>
		public bool IsDocked()
		{
			return (Docked);
		}
		/// <summary>
		/// Gets/Sets docking status based on events in schedule (not reliable after docking with revisions)
		/// </summary>
		private bool Docked
		{
			get { return (m_isDocked); }
			set { m_isDocked = value; }
		}
		/// <summary>
		/// Indicates if shuttle is in orbit
		/// </summary>
		private bool m_inOrbit;
		/// <summary>
		/// Has the shuttle landed?
		/// </summary>
		private bool m_landed;
		/// <summary>
		/// Getter/Setter for landing status
		/// </summary>
		private bool Landed
		{
			get { return(m_landed); }
			set { m_landed = value; }
		}
		/// <summary>
		/// Returns if the shuttle is in orbit or not
		/// </summary>
		/// <returns>true if the shuttle is in orbit (not reliable after launching with revisions)</returns>
		public bool InSpace()
		{
			return (InOrbit);
		}
		/// <summary>
		/// Gets/Sets In Orbit status
		/// </summary>
		private bool InOrbit
		{
			get
			{
				if (Landed)
					return (false);
				else
				{
					return (Orbit > 0.0);
				}
			}
			set { m_inOrbit = (value | (Orbit > 0.0)); }
		}

		/// <summary>
		/// Current date for the event
		/// </summary>
		private DateTime m_dtHeading;
		/// <summary>
		/// Gets/Sets the current date for the event
		/// </summary>
		private DateTime HeadingDate
		{
			get { return (m_dtHeading); }
			set
			{
				if (value != null)
					m_dtHeading = value;
			}
		}
		/// <summary>
		/// The number of orbits the shuttle has completed
		/// </summary>
		private double m_orbit;
		/// <summary>
		/// Gets/Sets the number of orbits for the shuttle
		/// </summary>
		private double Orbit
		{
			get { return (m_orbit); }
			set { m_orbit = value; }
		}
		/// <summary>
		/// The entry for the event
		/// </summary>
		private string m_subject;
		/// <summary>
		/// Gets/Sets the entry for the event
		/// </summary>
		private string Subject
		{
			get { return (m_subject); }
			set
			{
				if (value != null)
					m_subject = value.Trim();
				else
					m_subject = null;
			}
		}
		/// <summary>
		/// Site for the event
		/// </summary>
		private string m_site;
		/// <summary>
		/// Gets/Sets Site for the event
		/// </summary>
		private string Site
		{
			get { return (m_site); }
			set
			{
				if (value != null)
					m_site = value.Trim();
				else
					m_site = null;
			}
		}
		/// <summary>
		/// Mission Elapsed Time
		/// </summary>
		private string m_met;
		/// <summary>
		/// Gets/Sets Mission Elapsed Time
		/// </summary>
		private string MissionElapsedTime
		{
			get { return (m_met); }
			set
			{
				if (value != null)
					m_met = value.Trim();
				else
					m_met = null;
			}
		}
		/// <summary>
		/// Central Time
		/// </summary>
		private string m_central;
		/// <summary>
		/// Gets/Sets Central Time for the event
		/// </summary>
		private string CentralTime
		{
			get { return (m_central); }
			set
			{
				if (value != null)
					m_central = value.Trim();
				else
					m_central = null;
			}
		}
		/// <summary>
		/// Eastern Time
		/// </summary>
		private string m_eastern;
		/// <summary>
		/// Gets/Sets Eastern Time for the event
		/// </summary>
		private string EasternTime
		{
			get { return (m_eastern); }
			set { m_eastern = value.Trim(); }
		}
		/// <summary>
		/// GMT
		/// </summary>
		private string m_greenwich;
		/// <summary>
		/// Gets/Sets GMT
		/// </summary>
		private string GreenwichMeanTime
		{
			get { return (m_greenwich); }
			set
			{
				if (value != null)
					m_greenwich = value.Trim();
				else
					m_greenwich = null;
			}
		}
		/// <summary>
		/// Flight Day
		/// </summary>
		private string m_flightDay;
		/// <summary>
		/// Gets/Sets the Flight Day
		/// </summary>
		private string FlightDay
		{
			get { return (m_flightDay); }
			set
			{
				if (value != null)
					m_flightDay = value.Trim();
				else
					m_flightDay = null;
			}
		}
		/// <summary>
		/// Time Zone for the Viewer
		/// </summary>
		private string m_ViewingTimeZone;
		/// <summary>
		/// Gets/Sets the Time Zone for the Viewer
		/// </summary>
		public string ViewingTimeZone
		{
			get { return (m_ViewingTimeZone); }
			private set
			{
				if (value != null)
				{
					m_ViewingTimeZone = value.Trim();
					ViewingTimeZoneTZ = TimeZoneInfo.FindTimeZone(ViewingTimeZone);
				}
				else
				{
					m_ViewingTimeZone = null;
					ViewingTimeZoneTZ = null;
				}
			}
		}
		/// <summary>
		/// TimeZoneInfo of the Viewer
		/// </summary>
		private TimeZoneInfo m_tziViewingTimeZoneTZ;
		/// <summary>
		/// Gets/Sets TimeZoneInfo of the Viewer
		/// </summary>
		private TimeZoneInfo ViewingTimeZoneTZ
		{
			get { return (m_tziViewingTimeZoneTZ); }
			set
			{
				if (value != null)
					m_tziViewingTimeZoneTZ = value;
				else
					m_tziViewingTimeZoneTZ = null;
			}
		}
		/// <summary>
		/// Time Zone for Johnson Space Center
		/// </summary>
		private TimeZoneInfo m_tziJohnsonSpaceCenter;
		/// <summary>
		/// Gets/Sets Time Zone for Johnson Space Center
		/// </summary>
		private TimeZoneInfo JohnsonSpaceCenterTZ
		{
			get { return (m_tziJohnsonSpaceCenter); }
			set
			{
				if (value != null)
					m_tziJohnsonSpaceCenter = value;
				else
					m_tziJohnsonSpaceCenter = null;
			}
		}
		#endregion	//	PermanentVacations.Nasa.Sts.Schedule Variables with Getters/Setters

		#region Nasa Sts TV Constructors & supporting functions
		/// <summary>
		/// Initialize class with filename of Nasa TV Schedule
		/// </summary>
		/// <param name="excelFile">file name of Nasa Sts TV Schedule</param>
		public NasaStsTVSchedule(string excelFile)
		{
			PreInitializeClass();

			NasaTVScheduleFile = excelFile;

			PostInitializeExcel();
		}

		/// <summary>
		/// Initialize class with filename of Nasa TV Schedule & Viewing Time Zone
		/// </summary>
		/// <param name="excelFile">file name of Nasa Sts TV Schedule</param>
		/// <param name="viewingTimeZone">Time Zone for Viewer</param>
		public NasaStsTVSchedule(string excelFile, string viewingTimeZone)
		{
			PreInitializeClass();

			NasaTVScheduleFile = excelFile;

			ViewingTimeZone = viewingTimeZone;

			PostInitializeExcel();
		}

		/// <summary>
		/// Initialize class with filename of Nasa TV Schedule, Viewing Time Zone & method of interfacing with Excel
		/// Microsoft.Office.Interop.Excel is implemented (and the first to understand how to open a file)
		/// Microsoft.Office.Tools.Excel is not developed or working yet
		/// </summary>
		/// <param name="excelFile">file name of Nasa Sts TV Schedule</param>
		/// <param name="viewingTimeZone">Time Zone for Viewer</param>
		/// <param name="typeInterface">Type of Excel Interface (Interop or Tools)</param>
		public NasaStsTVSchedule(string excelFile, string viewingTimeZone, ExcelInterface typeInterface)
		{
			PreInitializeClass();

			NasaTVScheduleFile = excelFile;
			ViewingTimeZone = viewingTimeZone;
			ExcelTypeInterface = typeInterface;

			switch (typeInterface)
			{
				case ExcelInterface.InteropExcel:
				case ExcelInterface.ToolsExcel:
					ExcelTypeInterface = typeInterface;
					break;
				default:
					throw new ArgumentException(Properties.Resources.ERR_INVALID_ARGUMENT,
						Properties.Resources.ERR_ARGUMENT_TYPE_EXCEL);
			}
			PostInitializeExcel();
		}

		/// <summary>
		/// Initialize Excel Interface variables
		/// </summary>
		private void PostInitializeExcel()
		{
			//  ReadScheduleRow will make many calls to the convert time zone function ConvertTimeZoneToTimeZone
			TZInfo = Microsoft.Msdn.BclTeam.TimeZoneInfo.GetTimeZonesFromRegistry();
			JohnsonSpaceCenterTZ = TimeZoneInfo.FindTimeZone(Properties.Resources.TZ_JOHNSON_SPACE_CENTER);
			switch (ExcelTypeInterface)
			{
				case ExcelInterface.InteropExcel:
					InteropExcelIF = new InteropExcelInterface();
					break;
				case ExcelInterface.ToolsExcel:
					throw new NotImplementedException(Properties.Resources.ERR_MICROSOFT_OFFICE_TOOLS_EXCEL_NOT_SUPPORTED);
				default:
					throw new ArgumentException(Properties.Resources.ERR_INVALID_ARGUMENT,
						Properties.Resources.ERR_ARGUMENT_TYPE_EXCEL);
			}
		}

		/// <summary>
		/// Initialize common variables
		/// </summary>
		private void PreInitializeClass()
		{
			RowCount = 0;
			ExcelTypeInterface = ExcelInterface.InteropExcel;
			ViewingTimeZone = Properties.Resources.TZ_DEFAULT;

			ProcessingError = null;

			CurrentRow = 0;

			//	{ - Standard column numbers for NASA Schedule events
			OrbitColumnHeader = 1;
			SubjectColumnHeader = 3;
			SiteColumHeader = 4;
			FlightDayColumnHeader = 5;
			MissionElapsedTimeColumnHeader = 6;
			CentralTimeColumnHeader = 7;
			EasternTimeColumnHeader = 8;
			GreenwichMeanTimeColumnHeader = 9;
			//	} - Standard column numbers for NASA Schedule events

			DaylightSavingsTime = false;
			IsEOF = false;
			Docked = false;
			InOrbit = false;

			Orbit = 0;
			Subject = "";
			Site = "";
			MissionElapsedTime = "";
			CentralTime = "";
			EasternTime = "";
			GreenwichMeanTime = "";
			FlightDay = "";
			ViewingTimeZone = "";
			NasaTVScheduleFile = "";
		}

		/// <summary>
		/// Close the spreadsheet and Excel.
		/// 
		/// Do this before destroying the Excel application by setting it to null!
		/// </summary>
		public void Close()
		{
			switch (ExcelTypeInterface)
			{
				case ExcelInterface.InteropExcel:
					InteropExcelIF.Close();
					break;
				case ExcelInterface.ToolsExcel:
					break;
			}
		}
		#endregion	//	Nasa Sts TV Constructors & supporting functions

		#region Nasa Sts TV Destructor & supporting functions
		/// <summary>
		/// Class destructor for Nasa Sts TV Schedule decoder
		/// </summary>
		~NasaStsTVSchedule()
		{
			ReleaseResources();

			switch (ExcelTypeInterface)
			{
				case ExcelInterface.InteropExcel:
					ReleaseInteropExcelResources();
					break;
				case ExcelInterface.ToolsExcel:
					ReleaseToolsExcelResources();
					break;
				default:
					break;
			}
		}

		private void ReleaseResources()
		{
			rgDateHeader = null;
			rgEvaActivity = null;
			rgFlightDayHeader = null;
			rgIssCrewSleepActivity = null;
			rgShuttleCrewSleepActivity = null;
			MissionDurationTime = null;
			ProcessingError = null;
		}

		/// <summary>
		/// Release Microsoft.Office.Interop.Excel interface
		/// </summary>
		private void ReleaseInteropExcelResources()
		{
			InteropExcelIF = null;
		}

		/// <summary>
		/// Release Microsoft.Office.Tools.Excel interface
		/// </summary>
		private void ReleaseToolsExcelResources()
		{
			if (ToolsExcelIF != null)
			{
				ToolsExcelIF = null;
			}
		}
		#endregion	//	Nasa Sts TV Destructor & supporting functions

		#region Nasa Sts TV Schedule - Open, Read and Decode file
		/// <summary>
		/// Open Nasa TV Schedule (Excel file)
		/// </summary>
		private void OpenNasaTvSchedule()
		{
			RowCount = 0;
			ColumnCount = 0;
			CurrentRow = 0;
			try
			{
				switch (ExcelTypeInterface)
				{
					case ExcelInterface.InteropExcel:
						TvScheduleCells = InteropExcelIF.OpenExcelFile(NasaTVScheduleFile);
						break;
					case ExcelInterface.ToolsExcel:
						TvScheduleCells = ToolsExcelIF.OpenExcelFile(NasaTVScheduleFile);
						break;
					default:
						throw new ArgumentException(Properties.Resources.ERR_INVALID_ARGUMENT,
							Properties.Resources.ERR_ARGUMENT_TYPE_EXCEL);
				}
				RowCount = TvScheduleCells.GetUpperBound(0);
				ColumnCount = TvScheduleCells.GetUpperBound(1);
				CurrentRow = 1;
			}
			catch (COMException comExp)
			{
				if (Properties.Settings.Default.CopyExceptionsToClipboard)
					Clipboard.SetText(comExp.Message + CRLF + comExp.StackTrace, TextDataFormat.Text);
				throw;
			}
		}

		/// <summary>
		/// Read MASA TV Schedule
		/// Could generate an InvalidFileFormatException
		/// </summary>
		/// <returns>NasaStsTVScheduleEntry of scheduling information for event</returns>
		public NasaStsTVScheduleEntry ReadScheduleRow()
		{
			ScheduleType entryType = ScheduleType.empty;
			NasaStsTVScheduleEntry dataRow = null;
			if (!SuccessfullyOpened)
			{
				try
				{
					OpenNasaTvSchedule();
				}
				catch (InvalidFileFormatException invalidFile)
				{
					NasaStsTVScheduleEntry error = new NasaStsTVScheduleEntry(DateTime.MinValue, DateTime.MaxValue, false,
						invalidFile.Message, 0, invalidFile.StackTrace, "", ScheduleType.error);
					return (error);
				}
			}
			if (SuccessfullyOpened)
			{
				for (; !EOF() && (entryType != ScheduleType.scheduleEntry)
					&& (entryType != ScheduleType.error); CurrentRow++)
				{
					//	Could get an InvalidFileFormatException exception
					try
					{
						entryType = DecodeScheduleRow(CurrentRow);
					}
					catch (InvalidFileFormatException expInvalidFileFormat)
					{
						ProcessingError = expInvalidFileFormat;
						entryType = ScheduleType.error;
					}
				}
			}
			if (entryType == ScheduleType.scheduleEntry)
			{
				CurrentRow--;   //  CurrentRow is incremented before testing the return type of DecodeScheduleCurrentRow()
				dataRow = ProcessEntry(CurrentRow);
				if (dataRow == null)
				{
					entryType = ScheduleType.empty;
				}
				CurrentRow++;
			}
			else if (entryType == ScheduleType.error)
			{
				dataRow = new NasaStsTVScheduleEntry(DateTime.MinValue, DateTime.MinValue, false,
					ProcessingError.Message, 0, "InvalidFileFormatException", "", ScheduleType.error);
			}

			return (dataRow);
		}

		/// <summary>
		/// Decode entries in Nasa TV Schedule Excel spreadsheet
		/// Could generate an InvalidFileFormatException
		/// </summary>
		/// <param name="row">Row for the event to decode</param>
		/// <returns>Type of event</returns>
		private ScheduleType DecodeScheduleRow(int row)
		{
			ScheduleType typeEntry = ScheduleType.empty;

			object cellOrbit;

			if (CurrentRow < RowCount)
			{
				//	Year has not been initialized yet
				//  A revision or creation date is required in the spreadsheet before any headers are processed
				//  The revision/creation date is in the first few lines of the spreadsheet
				if (Year == 0)
				{
					GetCreationRevisionDate();
				}
				cellOrbit = TvScheduleCells.GetValue(row, OrbitColumnHeader);
				if (cellOrbit != null)
				{
					try
					{
						typeEntry = ProcessCellOrbit(cellOrbit, row);
					}
					catch (InvalidFileFormatException)
					{
						typeEntry = ScheduleType.error;
					}
				}
				else
				{
					if (IsRowScheduleEntry(row))
						typeEntry = ScheduleType.scheduleEntry;
				}
			}
			else
			{
				IsEOF = true;
			}

			return (typeEntry);
		}

		/// <summary>
		/// Sets the Year from the Creation or Revision date
		/// It is in the first few rows of the NASA banner in among
		/// the mission and how to receive NASA TV
		/// </summary>
		private void GetCreationRevisionDate()
		{
			Object cellOne = TvScheduleCells.GetValue(CurrentRow, 1);
			if ((cellOne != null) && (cellOne.GetType().ToString() == Properties.Resources.SYSTEM_STRING))
			{
				//  The Regular Expression to get the revision date is only accessed once, so it doesn't need to be compiled
				Regex rgDate = new Regex(Properties.Resources.RGX_MM_DD_YY);

				Match mtchDate = rgDate.Match(cellOne.ToString());

				GroupCollection grpcolDate = mtchDate.Groups;

				if (grpcolDate[Properties.Resources.IX_MONTH].Success &&
					grpcolDate[Properties.Resources.IX_DAY].Success &&
				grpcolDate[Properties.Resources.IX_YEAR].Success)
				{
					string revisionDate = cellOne.ToString().Trim();
					DateTime dtRevisionDate = DateTime.ParseExact(revisionDate,
						Properties.Resources.NASA_MM_DD_YY, CultureInfo.InvariantCulture);
					Year = dtRevisionDate.Year;
					if (Year < 2000)
						Year += 2000;
				}
				rgDate = null;
			}
		}

		/// <summary>
		/// Creates a NasaStsTVScheduleEntry for schedule entries
		/// </summary>
		/// <param name="row">Row for the schedule to capture</param>
		/// <returns>Event Schedule</returns>
		private NasaStsTVScheduleEntry ProcessEntry(int row)
		{
			//  Running into problems converting between timezones  .Net does not have the capability
			//  TimeZone information is local time

			DateTime dtCentral = HeadingDate;
			DateTime dtBeginViewingTime;
			DateTime dtEndViewingTime = HeadingDate;

			Changed = false;    //  Column 2 will contain an asterisk if an item has changed

			bool validEntry = false;

			NasaStsTVScheduleEntry entryRow = null;

			//  If there has been a flight missionDay heading process
			if (HeadingDate != DateTime.MinValue)
			{
				object cellTwo = TvScheduleCells.GetValue(row, 2);
				if (cellTwo != null)
				{
					System.Type cellTwoType = cellTwo.GetType();
					if (cellTwoType.FullName == Properties.Resources.SYSTEM_STRING)
					{
						string cellTwoValue = cellTwo.ToString();
						Changed = (cellTwoValue == Properties.Resources.NASA_CHANGED);
					}
				}
				Subject = GetMultiLineSubject(row);

				//  State variables for docking and in space are not reliable for schedule revisions
				//	published after launch or docking
				if (Subject.Contains(Properties.Resources.NASA_DOCKING))
				{
					if (!Subject.Contains(Properties.Resources.NASA_VTR_PLAYBACK))
					{
						Docked = true;
					}
				}
				else if (Subject.Contains(Properties.Resources.NASA_UNDOCKING) ||
				Subject.Contains(Properties.Resources.NASA_UNDOCKS))
				{
					if (!Subject.Contains(Properties.Resources.NASA_VTR_PLAYBACK))
					{
						Docked = false;
					}
				}
				else if (Subject == Properties.Resources.NASA_LAUNCH)
				{
					if (!Subject.Contains(Properties.Resources.NASA_VTR_PLAYBACK))
					{
						InOrbit = true;
					}
				}
				else if (Subject.Contains(Properties.Resources.NASA_LANDING))
				{
					if (!Subject.Contains(Properties.Resources.NASA_VTR_PLAYBACK))
					{
						InOrbit = false;
						Landed = true;
					}
				}
				//	20071226 - Ralph Hightower
				//		STS-118 schedule has a site entry of " " which none the less caused a blank site
				//		to be entered in the schedule (leading and trailing spaces are trimmed)
				bool nullSite = true;
				if (TvScheduleCells.GetValue(row, SiteColumHeader) != null)
				{
					Site = TvScheduleCells.GetValue(row, SiteColumHeader).ToString();
					Site = Site.Trim();
					nullSite = false;
				}
				if (nullSite || (Site.Length == 0))
				{
					if (Docked)
						Site = Properties.Resources.NASA_ISS;
					else
						Site = Properties.Resources.NASA_STS;

					if (Subject.Contains(Properties.Resources.NASA_CREW_SLEEP_BEGINS) ||
						Subject.Contains(Properties.Resources.NASA_CREW_WAKE_UP) ||
                        Subject.Contains(Properties.Resources.NASA_CREW_WAKEUP))
					{
						if (ISSCrewSleep(row) || ISSCrewWakeUp(row))
							Site = Properties.Resources.NASA_ISS;
						if (ShuttleCrewSleep(row) || ShuttleCrewWakeUp(row))
							Site = Properties.Resources.NASA_STS;
					}
				}
				if (TvScheduleCells.GetValue(row, MissionElapsedTimeColumnHeader) != null)
				{
					MissionElapsedTime = TvScheduleCells.GetValue(row, MissionElapsedTimeColumnHeader).ToString();
					MissionDurationTime.Set(TvScheduleCells, row, MissionElapsedTimeColumnHeader);
				}
				//  watch for "NET L" usually means "Net Landing + some time"
				if (TvScheduleCells.GetValue(row, OrbitColumnHeader) != null)
					Orbit = (System.Double)TvScheduleCells.GetValue(row, OrbitColumnHeader);
				if (TvScheduleCells.GetValue(row, FlightDayColumnHeader) != null)
					FlightDay = TvScheduleCells.GetValue(row, FlightDayColumnHeader).ToString();
				if (TvScheduleCells.GetValue(row, CentralTimeColumnHeader) != null)
				{
					CentralTime = ExcelFormatTime(row, CentralTimeColumnHeader);
					dtBeginViewingTime = ConvertFromCentralTzToViewingTz(dtCentral, CentralTime);
					dtEndViewingTime = GuesstimateFixedEvents(Subject, dtBeginViewingTime);
					//  If a special event was not found, get the start time for the next event
					if (dtEndViewingTime == DateTime.MinValue)
						dtEndViewingTime = ReadAhead();
					validEntry = true;
					//  This situation may not occur (except for STS Landing since though there are next events,
					//	the events are Net Landing + a time span
					//  If the end time occurs before the beginning time, assume 30 minutes
					if (dtBeginViewingTime > dtEndViewingTime)
						dtEndViewingTime = dtBeginViewingTime.AddHours(1);
					entryRow = new NasaStsTVScheduleEntry(dtBeginViewingTime, dtEndViewingTime, Changed, Subject,
						Orbit, Site, FlightDay, ScheduleType.scheduleEntry);
				}
			}

			if (validEntry)
			{
				return (entryRow);
			}
			else
				return (null);
		}

		/// <summary>
		/// Some NASA STS TV Schedules have the event on multiple lines instead in a single cellTime
		/// This routine reads ahead concatenating the lines until an empty cellTime is found or until the end of the spreadsheet is found
		/// </summary>
		/// <param name="row">Row to begin building the subject line</param>
		/// <returns>Concatenated Subject or a single line entry</returns>
		private string GetMultiLineSubject(int row)
		{
			StringBuilder bldrMultiLineSubject = new StringBuilder();

			int nextRow = row;

			Object singleLine = TvScheduleCells.GetValue(nextRow, SubjectColumnHeader);
			while ((nextRow < RowCount) && (singleLine != null) &&
				(singleLine.GetType().FullName == Properties.Resources.SYSTEM_STRING) &&
				singleLine.ToString() != Properties.Resources.NASA_DEFINITION_OF_TERMS)
			{
				if (singleLine.ToString() != Properties.Resources.NASA_DEFINITION_OF_TERMS)
				{
					bldrMultiLineSubject.Append(singleLine.ToString().Trim());
					bldrMultiLineSubject.Append(" ");
				}
				nextRow++;
				singleLine = TvScheduleCells.GetValue(nextRow, SubjectColumnHeader);
			}

			string multiLineSubject = bldrMultiLineSubject.ToString().Trim();

			bldrMultiLineSubject = null;
			return (multiLineSubject);
		}

		/// <summary>
		/// Some events do not last until the next event occurs, such as press conferences, PAO events, etc
		/// This uses the EventTimes lookup table to get approximate times of duration
		/// </summary>
		/// <param name="subject">Scheduled Event</param>
		/// <param name="beginTime">Start of Event</param>
		/// <returns>End of Event, if found, otherwise DateTime.MinTime</returns>
		private DateTime GuesstimateFixedEvents(string subject, DateTime beginTime)
		{
			DateTime dtEndViewing = DateTime.MinValue;

			int maxTimedEvents = EventTimes.GetUpperBound(0);
			//  Some events do not last until the next event, such as flight missionDay highlights, press conferences,
			//	interviews, etc. This uses some rough estimates for those events.
			//  The guesstimate for the event times are recorded in the resources file as comments
			//  A TO DO for me (Ralph Hightower) is to create a method to read from the resource file and get the times
			//	from the comments.  Expediency ruled in favor of duplicating the information in a two dimensional string array
			int indexFixedEvents;
			for (indexFixedEvents = EventTimes.GetLowerBound(0); indexFixedEvents <= maxTimedEvents; indexFixedEvents++)
			{
				//  The Flight Day Highlight event is a regular expression since each missionDay is a new number
				if (EventTimes[indexFixedEvents, 0] == Properties.Resources.TM_RG_FLIGHT_DAY_HIGHLIGHTS)
				{
					Match mFlightDayHighlights = rgFlightDayHighlights.Match(Subject);
					GroupCollection grpcolFlightDayHighlights = mFlightDayHighlights.Groups;

					if (grpcolFlightDayHighlights[Properties.Resources.IX_FLIGHT_DAY].Success &&
						grpcolFlightDayHighlights[Properties.Resources.IX_DAY].Success &&
						grpcolFlightDayHighlights[Properties.Resources.IX_HIGHLIGHTS].Success)
					{
						TimeSpan dtApproximateTime = TimeSpan.Parse(EventTimes[indexFixedEvents, 1]);
						dtEndViewing = beginTime.Add(dtApproximateTime);
					}
				}
				else if (Subject.Contains(EventTimes[indexFixedEvents, 0]))
				{
					TimeSpan dtApproximateTime = TimeSpan.Parse(EventTimes[indexFixedEvents, 1]);
					dtEndViewing = beginTime.Add(dtApproximateTime);
					break;
				}
			}
			return (dtEndViewing);
		}

		/// <summary>
		/// Looks ahead to find the next event
		/// Mini schedule decoder to set the end time of the current entry to the start time of the next entry
		///
		/// returns DateTime of start time in Central Time of the next entry
		/// </summary>
		/// <returns>Date and Time for start of next event, which provides the ending date and time for the current event</returns>
		private DateTime ReadAhead()
		{
			//  Initialize End Date to the Begin Date
			//  The for loop will capture the End Date if a Date Heading is process before a schedule entry
			DateTime endDate = HeadingDate;
			int indexRowAhead;
			string endCentralTime = "";
			ScheduleType scheduleRow = ScheduleType.empty;
			bool tempEOF = false;

			bool shuttleCrewSleep = false;
			bool issCrewSleep = false;
			bool evaBegins = false;

			string subjectStart = TvScheduleCells.GetValue(CurrentRow, SubjectColumnHeader).ToString();
			if (subjectStart.Contains(Properties.Resources.NASA_EVA) &&
				(subjectStart.Contains(Properties.Resources.NASA_BEGINS)))
			{
				evaBegins = EVABegins(CurrentRow);
			}
			if (subjectStart.Contains(Properties.Resources.NASA_CREW_SLEEP_BEGINS))
			{
				shuttleCrewSleep = ShuttleCrewSleep(CurrentRow);
				issCrewSleep = ISSCrewSleep(CurrentRow);
			}

			for (indexRowAhead = CurrentRow + 1; (indexRowAhead <= RowCount) &&
				(scheduleRow != ScheduleType.scheduleEntry) &&
				!tempEOF && !ErrorProcessed; indexRowAhead++)
			{
				if (IsRowScheduleEntry(indexRowAhead))
				{
					scheduleRow = ScheduleType.scheduleEntry;
					endCentralTime = ExcelFormatTime(indexRowAhead, CentralTimeColumnHeader);
					endDate = ConvertFromCentralTzToViewingTz(endDate, endCentralTime);
				}
				else
				{
					object cellOrbit = TvScheduleCells.GetValue(indexRowAhead, OrbitColumnHeader);
					if (cellOrbit != null)
					{
						System.Type cellOrbitType = cellOrbit.GetType();
						if (cellOrbitType.FullName == Properties.Resources.SYSTEM_STRING)
						{
							string endDateHeader = cellOrbit.ToString();
							//  Don't loop through missionDay for end of file record
							if (endDateHeader.Contains(Properties.Resources.NASA_DEFINITION_OF_TERMS))
							{
								tempEOF = true;
								break;
							}
							//  Don't loop through missionDay if a Header Record
							if (endDateHeader == Properties.Resources.NASA_ORBIT)
								continue;
							//  Don't loop through missionDay if a Flight Day Record
							if ((endDateHeader.Length > 2) && (endDateHeader.Substring(0, 2) == Properties.Resources.NASA_FD))
								continue;
							//  This is probably a Date Header Record
							if (MatchDateHeader(endDateHeader))
							{
								try
								{
									endDate = ProcessDateHeader(endDateHeader, false);
								}
								catch (InvalidFileFormatException invalidFileFormat)
								{
									ProcessingError = invalidFileFormat;
									break;
								}
								continue;
							}
						}
					}
				}
				//  If the Shuttle Crew or ISS Crew is in a sleep period, look for their wake up entry
				if ((shuttleCrewSleep || issCrewSleep) && (scheduleRow == ScheduleType.scheduleEntry))
				{
                    string strSubject = TvScheduleCells.GetValue(indexRowAhead, SubjectColumnHeader).ToString();
					if (strSubject.Contains(Properties.Resources.NASA_CREW_WAKE_UP) ||
						strSubject.Contains(Properties.Resources.NASA_CREW_WAKEUP))
					{
						if (shuttleCrewSleep && ShuttleCrewWakeUp(indexRowAhead))
						{
							scheduleRow = ScheduleType.scheduleEntry;
							break;
						}
						else
							scheduleRow = ScheduleType.empty;
						if (issCrewSleep && ISSCrewWakeUp(indexRowAhead) && !shuttleCrewSleep)
						{
							scheduleRow = ScheduleType.scheduleEntry;
							break;
						}
						else
							scheduleRow = ScheduleType.empty;

					}
					else
						scheduleRow = ScheduleType.empty;
				}
				if (evaBegins && (scheduleRow == ScheduleType.scheduleEntry))
				{
					if (TvScheduleCells.GetValue(indexRowAhead, SubjectColumnHeader).ToString().Contains(Properties.Resources.NASA_ENDS))
					{
						if (!EVAEnds(indexRowAhead))
						{
							scheduleRow = ScheduleType.empty;
						}
					}
					else
						scheduleRow = ScheduleType.empty;
				}
			}
			return (endDate);
		}

		/// <summary>
		/// Process schedule entry based on the content in column 1
		/// This can have many different formats
		/// 1. Comments
		/// 2. Header Record (ORBIT, SUBJECT, SITE, MET, C[SD]T, E[SD]T, GMT
		/// 3. Date Header (DAYOFWEEK, MONTH Day)
		/// 4. Flight Day Header (FD \d*)
		/// 5. Definitions (not processed)
		/// </summary>
		/// <param name="cellOrbit">Cell Value for Orbit column</param>
		/// <param name="row">Current Row</param>
		/// <returns>Type of schedule for the current row</returns>
		private ScheduleType ProcessCellOrbit(object cellOrbit, int row)
		{
			ScheduleType typeEntry = ScheduleType.empty;
			System.Type cellOrbitType = cellOrbit.GetType();
			switch (cellOrbitType.FullName)
			{
				case "System.String":
					{
						string cellOrbitValue = (string)cellOrbit.ToString();
						//  Row contains "DEFINITION OF TERMS" which is the end of file; no schedule entries
						//	exist after this value.  What remains are the definitions of the acronyms used in the schedule
						if (cellOrbitValue.Contains(Properties.Resources.NASA_DEFINITION_OF_TERMS))
						{
							IsEOF = true;
							typeEntry = ScheduleType.definitionOfTerms;
						}
						//  Header: ORBIT(1)   SUBJECT(3) SITE(4)    MET(6) C[SD]T(7)  E[SD]T(8)  GMT(9)
						//  Cell number in parenthesis
						else if (cellOrbitValue == Properties.Resources.NASA_ORBIT)
						{
							typeEntry = ScheduleType.columnHeading;
							ProcessOrbitHeader(row);
						}
						else
						{
							if (MatchDateHeader(cellOrbitValue))
							{
								try
								{
									HeadingDate = ProcessDateHeader(cellOrbitValue, true);
									typeEntry = ScheduleType.dateHeading;
								}
								catch (InvalidFileFormatException expInvalidFileFormat)
								{
									ProcessingError = expInvalidFileFormat;
								}
								finally
								{
									if (ProcessingError != null)
										typeEntry = ScheduleType.error;
								}
								break;  //  Have the Date, no need to check any other missionDay
							}

							//  if a date heading wasn't found, look for Flight Day heading (FD \d .*/ FD \d)
							if (MatchFlightDayHeader(cellOrbitValue))
							{
								typeEntry = ScheduleType.flightDayHeading;
							}
						}
					}
					break;
				//  Row containing orbit value must have a entry and central time, besides mission elapsed time and eastern time
				case "System.Double":
					{
						//  Know that Orbit column has a number
						//  Do the columns, Subject, Central Time, Eastern Time, and GMT contain String, Double, Double, Double?
						if (IsRowScheduleEntry(row))
							typeEntry = ScheduleType.scheduleEntry;
					}
					break;
				default:
					typeEntry = ScheduleType.empty;
					break;
			}
			return (typeEntry);
		}

		/// <summary>
		/// Matches FD #d
		/// </summary>
		/// <param name="cellTime"></param>
		/// <returns>true if header record is a Flight Day Header</returns>
		private bool MatchFlightDayHeader(string cell)
		{
			Match mFD = rgFlightDayHeader.Match(cell);
			GroupCollection grpcolFD = mFD.Groups;
			bool flightDayHeader = (grpcolFD[Properties.Resources.IX_FLIGHT_DAY].Success &&
				grpcolFD[Properties.Resources.IX_DAY].Success);
			return (flightDayHeader);
		}

		/// <summary>
		/// Determines if schedule entry based on the system type of the cells
		/// C3 (Subject) is a string
		/// C7 (Central Time) is a double
		/// C8 (Easter Time) is a double
		/// C9 (GMT) is a double
		/// </summary>
		/// <param name="row">Row number of schedule to examine</param>
		/// <returns>true if the row matches the criteria</returns>
		private bool IsRowScheduleEntry(int row)
		{
			bool scheduleEntry = false;
			object cellSubject = TvScheduleCells.GetValue(row, SubjectColumnHeader);
			object cellCentral = TvScheduleCells.GetValue(row, CentralTimeColumnHeader);
			object cellEastern = TvScheduleCells.GetValue(row, EasternTimeColumnHeader);
			object cellGreenwich = TvScheduleCells.GetValue(row, GreenwichMeanTimeColumnHeader);
			//  To be a schedule entry, there must be a Subject, Central Time, Eastern Time, and GMT
			//  Orbit and Mission Elapsed Time are optional
			if ((cellSubject != null) && (cellCentral != null) && (cellEastern != null) && (cellGreenwich != null))
			{
				System.Type cellSubjecType = cellSubject.GetType();
				System.Type cellCentralType = cellCentral.GetType();
				System.Type cellEasternType = cellEastern.GetType();
				System.Type cellGreenwichType = cellGreenwich.GetType();
				if ((cellSubjecType.FullName == Properties.Resources.SYSTEM_STRING)
					&& (cellCentralType.FullName == Properties.Resources.SYSTEM_DOUBLE)
					&& (cellEasternType.FullName == Properties.Resources.SYSTEM_DOUBLE)
					&& (cellGreenwichType.FullName == Properties.Resources.SYSTEM_DOUBLE))
					scheduleEntry = true;
			}
			return (scheduleEntry);
		}

		/// <summary>
		/// Captures the column where the Column Headers
		/// </summary>
		/// <param name="row">Current Row</param>
		/// <returns>true</returns>
		private bool ProcessOrbitHeader(int row)
		{
			int col;
			int headingCount = 0;
			//Header:  ORBIT   SUBJECT SITE    MET C[DS]T  E[DS]T  GMT
			for (col = 1; col <= ColumnCount; col++)
			{
				if (TvScheduleCells.GetValue(row, col) != null)
				{
					if (TvScheduleCells.GetValue(row, col).ToString() == Properties.Resources.NASA_ORBIT)
					{
						OrbitColumnHeader = col;
						headingCount++;
					}
					else if (TvScheduleCells.GetValue(row, col).ToString() == Properties.Resources.NASA_SUBJECT)
					{
						SubjectColumnHeader = col;
						headingCount++;
					}
					else if (TvScheduleCells.GetValue(row, col).ToString() == Properties.Resources.NASA_SITE)
					{
						SiteColumHeader = col;
						headingCount++;
					}
					else if (TvScheduleCells.GetValue(row, col).ToString() == Properties.Resources.NASA_MET)
					{
						MissionElapsedTimeColumnHeader = col;
						headingCount++;
					}
					else if ((TvScheduleCells.GetValue(row, col).ToString() == Properties.Resources.NASA_CDT)
					|| (TvScheduleCells.GetValue(row, col).ToString() == Properties.Resources.NASA_CST))
					{
						CentralTimeColumnHeader = col;
						DaylightSavingsTime = TvScheduleCells.GetValue(row, col).ToString().Substring(1, 1) == "D";
						headingCount++;
					}
					else if ((TvScheduleCells.GetValue(row, col).ToString() == Properties.Resources.NASA_EDT)
					|| (TvScheduleCells.GetValue(row, col).ToString() == Properties.Resources.NASA_EST))
					{
						EasternTimeColumnHeader = col;
						DaylightSavingsTime = TvScheduleCells.GetValue(row, col).ToString().Substring(1, 1) == "D";
						headingCount++;
					}
					else if (TvScheduleCells.GetValue(row, col).ToString() == Properties.Resources.NASA_GMT)
					{
						GreenwichMeanTimeColumnHeader = col;
						headingCount++;
					}
				}
			}
			if (headingCount < 7)
			{
				string InvalidOrbitHeader = Properties.Resources.ERR_INVALID_COLUMN_HEADER +
					row.ToString(CultureInfo.CurrentCulture);
				throw new ApplicationException(InvalidOrbitHeader);
			}
			return (true);
		}

		/// <summary>
		/// Gets the current date from the Flight Day heading
		/// Converts to DateTime
		/// </summary>
		/// <param name="weekdayMonthDay">Cell Value of the Date Header</param>
		/// <returns>DateTime of the Cell Value</returns>
		private DateTime ProcessDateHeader(string weekdayMonthDay, bool advanceYear)
		{
			DateTime dtHeading = DateTime.MinValue;
			bool dateHeaderFound = MatchDateHeader(weekdayMonthDay);
			if (dateHeaderFound)
			{
				int holdYear = Year;
				//  Date Heaader is DayOfWeek, Month missionDay in uppercase
				Match mtchDateHeader = rgDateHeader.Match(weekdayMonthDay);

				GroupCollection grpcolDateHeader = mtchDateHeader.Groups;

				string month = grpcolDateHeader[Properties.Resources.IX_MONTH].Value;
				int indexMonth;
				for (indexMonth = Months.GetLowerBound(0); indexMonth <= Months.GetUpperBound(0) &&
				month != Months[indexMonth]; indexMonth++)
				{
				}
				indexMonth++;	//	Month is zero based

				//	This could fire InvalidFileFormatException if the spreadsheet does not have a
				//	Create/Revision date that supplies the year of the mission
				Month = indexMonth;

				Day = Convert.ToInt32(grpcolDateHeader[Properties.Resources.IX_DAY].Value, CultureInfo.InvariantCulture);
				dtHeading = new DateTime(Year, Month, Day);
				if (!advanceYear)
					Year = holdYear;
			}

			return (dtHeading);
		}

		/// <summary>
		/// Determines if the current row is a Date Header (Weekday, Month, Day)
		/// </summary>
		/// <param name="weekdayMonthDay"></param>
		/// <returns>true if the current row is a Date Header</returns>
		private bool MatchDateHeader(string weekdayMonthDay)
		{
			Match mtchDateHeader = rgDateHeader.Match(weekdayMonthDay);

			GroupCollection grpcolDateHeader = mtchDateHeader.Groups;

			bool dateHeaderFound = grpcolDateHeader[Properties.Resources.IX_DAY_OF_WEEK].Success &&
				grpcolDateHeader[Properties.Resources.IX_MONTH].Success &&
				grpcolDateHeader[Properties.Resources.IX_DAY].Success;

			return (dateHeaderFound);
		}

		/// <summary>
		/// Formats the time of the weekdayMonthDay according to Excel method (Interop or Excel)
		/// </summary>
		/// <param name="row">Row of spreadsheet</param>
		/// <param name="weekdayMonthDay">Column of spreadsheet</param>
		/// <returns>Time as a string formatted similar to DateTime.ToString("hh:mm tt")</returns>
		private string ExcelFormatTime(int row, int cell)
		{
			string formattedTime = "";
			switch (ExcelTypeInterface)
			{
				case ExcelInterface.InteropExcel:
					formattedTime = DateTime.FromOADate((double)TvScheduleCells.GetValue(row, cell)).ToString("hh:mm tt");
					break;
				case ExcelInterface.ToolsExcel:
					formattedTime = ToolsExcelIF.FormatTime(TvScheduleCells, row, cell);
					break;
				default:
					throw new ArgumentException(Properties.Resources.ERR_EXCEL_FORMAT_TIME,
						Properties.Resources.ERR_ARGUMENT_TYPE_EXCEL);
			}
			return (formattedTime);
		}

		/// <summary>
		/// The Nasa TV Schedule is Houston-centric.  This is an easy method to convert from Central to other time zones
		/// There is a kludge for that 2 AM hour that does not occur when Daylight Savings Time ends and Standard Time begins for Eastern Time
		///
		/// Uses TimeZoneInfo developed by Microsoft Msdn BCPLTeam
		/// </summary>
		/// <param name="dtConvert">Date of the event in Central Time Zone</param>
		/// <param name="timeOfday">Time of the event in Central Time Zone</param>
		/// <returns>DateTime in Viewer's Time Zone</returns>
		private DateTime ConvertFromCentralTzToViewingTz(DateTime dtConvert, string timeOfday)
		{
			string convertTime = timeOfday.Trim();
			DateTime dtCentralTZ = dtConvert.Date;
			//TimeSpan tsTimeOfDay = TimeSpan.Parse(convertTime);
			DateTime dtTimeOfDay = DateTime.Parse(convertTime, CultureInfo.CurrentCulture);
			dtCentralTZ = dtCentralTZ.Add(dtTimeOfDay.TimeOfDay);
			DateTime dtViewingTZ = TimeZoneInfo.ConvertTimeZoneToTimeZone(dtCentralTZ, JohnsonSpaceCenterTZ, ViewingTimeZoneTZ);
			// Kludge for Eastern Daylight Time transition to Eastern Standard Time
			if ((dtCentralTZ.Hour == dtViewingTZ.Hour) && (ViewingTimeZoneTZ.DisplayName == Properties.Resources.TZ_US_EASTERN))
				dtViewingTZ = dtViewingTZ.AddHours(1);

			return (dtViewingTZ);
		}

		#endregion	//	Nasa Sts TV Schedule - Open, Read and Decode file

		#region Shuttle Crew Sleep Activity
		/// <summary>
		///  Regular expression built for C# on: Thu, Nov 8, 2007, 11:47:39 AM
		///  Using Expresso Version: 3.0.2766, http://www.ultrapico.com
		///
		/// (?<Shuttle>ATLANTIS|DISCOVERY|ENDEAVOR)
		/// (?:\s*/?\s*)?
		/// (?<ISS>ISS)?
		/// (?:\s*)
		/// (?<Activity>CREW SLEEP BEGINS|CREW WAKE UP)
		///
		///  Required: Shuttle (ATLANTIS, DISCOVERY, ENDEAVOR)
		///  Optional: ISS
		///  Required: Sleep Activity (CREW WAKE UP or SLEEP PERIOD BEGINS)
		///
		///  A description of the regular expression:
		///
		///  [Shuttle]: A named capture group. [ATLANTIS|DISCOVERY|ENDEAVOR]
		///      Select from 3 alternatives
		///          ATLANTIS
		///              ATLANTIS
		///          DISCOVERY
		///              DISCOVERY
		///          ENDEAVOR
		///              ENDEAVOR
		///  Match expression but don't capture it. [\s*/?\s*], zero or one repetitions
		///      \s*/?\s*
		///          Whitespace, any number of repetitions
		///          /, zero or one repetitions
		///          Whitespace, any number of repetitions
		///  [ISS]: A named capture group. [ISS], zero or one repetitions
		///      ISS
		///          ISS
		///  Match expression but don't capture it. [\s*]
		///      Whitespace, any number of repetitions
		///  [Action]: A named capture group. [CREW SLEEP BEGINS|CREW WAKE UP]
		///      Select from 2 alternatives
		///          CREW SLEEP BEGINS
		///              CREW
		///              Space
		///              SLEEP
		///              Space
		///              BEGINS
		///          CREW WAKE UP
		///              CREW
		///              Space
		///              WAKE
		///              Space
		///              UP
		///
		///
		/// </summary>

		/// <summary>
		/// Gets the status of Shuttle Crew Sleep
		/// </summary>
		/// <param name="row">Row of event to interpret</param>
		/// <returns>true if in sleep period</returns>
		private bool ShuttleCrewSleep(int row)
		{
			bool shuttleCrewSleep = SubjectVerbPatternMatch(rgShuttleCrewSleepActivity,
				Properties.Resources.IX_SHUTTLE, Properties.Resources.NASA_CREW_SLEEP_BEGINS, row);

			return (shuttleCrewSleep);
		}

		/// <summary>
		/// Gets the status of Shuttle Crew Wakeup
		/// </summary>
		/// <param name="row">Row of event to interpret</param>
		/// <returns>true if subject in wake up period</returns>
		private bool ShuttleCrewWakeUp(int row)
		{
			bool shuttleCrewWakeUp = SubjectVerbPatternMatch(rgShuttleCrewSleepActivity,
				Properties.Resources.IX_SHUTTLE, Properties.Resources.NASA_CREW_WAKE_UP, row);
            //  There is an entry in STS-122\tvshed_rev0.xls that has "CREW WAKEUP" instead of "CREW WAKE UP"
            if (!shuttleCrewWakeUp)
                shuttleCrewWakeUp = SubjectVerbPatternMatch(rgShuttleCrewSleepActivity,
                    Properties.Resources.IX_SHUTTLE, Properties.Resources.NASA_CREW_WAKEUP, row);

			return (shuttleCrewWakeUp);
		}
		#endregion	//	Shuttle Crew Sleep Activity

		#region ISS Crew Sleep Activity
		/// <summary>
		///  Regular expression built for C# on: Thu, Nov 8, 2007, 11:43:41 AM
		///  Using Expresso Version: 3.0.2766, http://www.ultrapico.com
		///
		/// (?<Shuttle>ATLANTIS|DISCOVERY|ENDEAVOR)?
		/// (?:\s*/?\s*)?
		/// (?<ISS>ISS)
		/// (?:\s*)
		/// (?<Activity>CREW SLEEP BEGINS|CREW WAKE UP)
		///
		///  Optional: name of Shuttle (ATLANTIS, DISCOVERY, or ENDEAVOR)
		///  Required: ISS
		///  Required: CREW WAKE UP or CREW SLEEP BEGINS
		///
		///  A description of the regular expression:
		///
		///  [Shuttle]: A named capture group. [ATLANTIS|DISCOVERY|ENDEAVOR], zero or one repetitions
		///      Select from 3 alternatives
		///          ATLANTIS
		///              ATLANTIS
		///          DISCOVERY
		///              DISCOVERY
		///          ENDEAVOR
		///              ENDEAVOR
		///  Match expression but don't capture it. [\s*/?\s*], zero or one repetitions
		///      \s*/?\s*
		///          Whitespace, any number of repetitions
		///          /, zero or one repetitions
		///          Whitespace, any number of repetitions
		///  [ISS]: A named capture group. [ISS]
		///      ISS
		///          ISS
		///  Match expression but don't capture it. [\s*]
		///      Whitespace, any number of repetitions
		///  [Action]: A named capture group. [CREW SLEEP BEGINS|CREW WAKE UP]
		///      Select from 2 alternatives
		///          CREW SLEEP BEGINS
		///              CREW
		///              Space
		///              SLEEP
		///              Space
		///              BEGINS
		///          CREW WAKE UP
		///              CREW
		///              Space
		///              WAKE
		///              Space
		///              UP
		/// </summary>

		/// <summary>
		/// Gets status of ISS subject sleep period
		/// </summary>
		/// <param name="row">Row of event to interpret</param>
		/// <returns>true if ISS subject in sleep period</returns>
		private bool ISSCrewSleep(int row)
		{
			bool issCrewSleep = SubjectVerbPatternMatch(rgIssCrewSleepActivity,
				Properties.Resources.IX_ISS, Properties.Resources.NASA_CREW_SLEEP_BEGINS, row);

			return (issCrewSleep);
		}

		/// <summary>
		/// Gets status of ISS subject wake up
		/// </summary>
		/// <param name="row">Row of event to interpret</param>
		/// <returns>true if ISS subject is in wake up period</returns>
		private bool ISSCrewWakeUp(int row)
		{
			bool issCrewWakeUp = SubjectVerbPatternMatch(rgIssCrewSleepActivity,
				Properties.Resources.IX_ISS, Properties.Resources.NASA_CREW_WAKE_UP, row);
            //  There is an entry in STS-122\tvshed_rev0.xls that has "CREW WAKEUP" instead of "CREW WAKE UP"
            if (!issCrewWakeUp)
                issCrewWakeUp = SubjectVerbPatternMatch(rgIssCrewSleepActivity,
                    Properties.Resources.IX_ISS, Properties.Resources.NASA_CREW_WAKEUP, row);

			return (issCrewWakeUp);
		}
		#endregion	//	ISS Crew Sleep Activity

		#region Versatile Regular Expression Pattern Match Subject with Verb
		/// <summary>
		/// Helper method used by:
		/// 1. ShuttleCrewSleepBegins
		/// 2. ShuttleCrewWakeup
		/// 3. ISSCrewSleepBegins
		/// 4. ISSCrewWakeUp
		/// 5. EVABegins
		/// 6. EVAEnds
		/// </summary>
		/// <param name="rgSubjectVerbPattern">Regular expression for required rgSubjectVerbPattern: Shuttle or ISS</param>
		/// <param name="subject">Crew: Shuttle or ISS</param>
		/// <param name="verb">CREW WAKE UP or CREW SLEEP BEGINS</param>
		/// <param name="row">Row in TvScheduleCells with Subject to match</param>
		/// <returns>true if Required Crew is in the desired Sleep or Wake Activity</returns>
		private bool SubjectVerbPatternMatch(Regex rgSubjectVerbPattern, string subject, string verb, int row)
		{
			string entry = TvScheduleCells.GetValue(row, SubjectColumnHeader).ToString(); ;

			Match mtchSubjectVerb = rgSubjectVerbPattern.Match(entry);

			GroupCollection grpcollSubjectVerb = mtchSubjectVerb.Groups;

			bool matchSubjectVerb = grpcollSubjectVerb[subject].Success &&
				(grpcollSubjectVerb[Properties.Resources.IX_ACTIVITY].Success &&
				(grpcollSubjectVerb[Properties.Resources.IX_ACTIVITY].ToString() == verb));

			return (matchSubjectVerb);
		}

		/// <summary>
		/// Checks to see if Subject is EVA BEGINS
		/// </summary>
		/// <param name="row">Row in TvScheduleCells with Subject to match</param>
		/// <returns>Returns true if the Subject contains EVA BEGINS</returns>
		private bool EVABegins(int row)
		{
			bool evaBegins = false;

			string entry = TvScheduleCells.GetValue(row, SubjectColumnHeader).ToString();
			//	Do not check the Subject-Verb Pattern match if the entry contains any of the
			//	EVA preparations to purge nitrogen from the bloodstream to avoid "the bends"
			if (!entry.Contains(Properties.Resources.NASA_CAMPOUT) &&
				!entry.Contains(Properties.Resources.NASA_PRE_BREATHE) &&
				!entry.Contains(Properties.Resources.NASA_PREBREATHE))
			{
				evaBegins = SubjectVerbPatternMatch(rgEvaActivity, Properties.Resources.IX_EVA,
					Properties.Resources.NASA_BEGINS, row);
			}

			return (evaBegins);
		}

		/// <summary>
		/// Checks to see if Subject is EVA ENDS
		/// </summary>
		/// <param name="row">Row in TvScheduleCells with Subject to match</param>
		/// <returns>Returns true if the Subject contains EVA ENDS</returns>
		private bool EVAEnds(int row)
		{
			bool evaEnds = SubjectVerbPatternMatch(rgEvaActivity, Properties.Resources.IX_EVA,
				Properties.Resources.NASA_ENDS, row);

			return (evaEnds);
		}
		#endregion	//	Versatile Regular Expression Pattern Match Subject with Verb
	}
	#endregion

	#region NasaStsTVSchedule Exceptions
	/// <summary>
	/// InvalidFileFormatException
	///		Possible causes:
	///		1) No Creation or Revision Date (to get the Year) in first couple of rows before a Date Header is reached
	///		2) No Print_Area range defined in the spreadsheet (COMException caught when opening spreadsheet)
	/// </summary>
	public class InvalidFileFormatException : ApplicationException
	{
		/// <summary>
		/// Text of InvalidFileFormatException
		/// </summary>
		static string invalidFileFormat = Properties.Resources.EXP_INVALID_FILE_FORMAT;

		/// <summary>
		/// Constructor for new InvalidFileFormatException
		/// </summary>
		/// <param name="auxMessage"></param>
		public InvalidFileFormatException(string auxMessage)
			: base(String.Format("{0} - {1}", invalidFileFormat, auxMessage))
		{
			this.Source = Properties.Resources.EXP_SOURCE;
		}

		/// <summary>
		/// Constructor for InvalidFileFormatException (another exception caught)
		/// </summary>
		/// <param name="auxMessage"></param>
		/// <param name="inner"></param>
		public InvalidFileFormatException(string auxMessage, Exception inner)
			: base(String.Format("{0} - {1}", invalidFileFormat, auxMessage), inner)
		{
		}
	}
	#endregion	//	NasaStsTVSchedule Exceptions

	#region MissionDuration - Handles Mission Elapsed Time
	/// <summary>
	/// Class definition for Mission Elapsed Time
	/// Contains Day, Time
	/// </summary>
	class MissionDuration : Object
	{
		/// <summary>
		/// TimeSpan for Mission Elapsed Time
		///		uses Day + TimeOfDay for duration
		/// </summary>
		private TimeSpan time;
		/// <summary>
		/// Compiled Regular Expression to extract the Mission Day from the column before the Mission Elapsed Time
		///		Uses Group Collection Index: IX_DAY
		/// </summary>
		private Regex crgxMissionDay;
		/// <summary>
		/// Getter/Setter for Regular Expression of the Mission Day
		/// </summary>
		private Regex rgxMissionDay
		{
			get
			{
				if (crgxMissionDay == null)
					crgxMissionDay = new Regex(Properties.Resources.RGX_MISSIONDAY, RegexOptions.Compiled);
				return (crgxMissionDay);
			}
			set
			{
				if (value == null)
					crgxMissionDay = null;
				else
					crgxMissionDay = value;
			}
		}

		/// <summary>
		/// Constructor for the Mission Duration class
		/// </summary>
		public MissionDuration()
		{
			time = TimeSpan.MinValue;
		}

		/// <summary>
		/// Destructor for the Mission Duration class
		///		Releases resources created
		/// </summary>
		~MissionDuration()
		{
			ReleaseResources();
		}

		/// <summary>
		/// Releases objects created
		/// </summary>
		private void ReleaseResources()
		{
			rgxMissionDay = null;
		}

		/// <summary>
		/// Returns the Mission Elapsed Time
		/// </summary>
		/// <returns>Mission Elapsed Time in Day/Hour:Mminue (Day, Hour, Minute in 2 digits each)</returns>
		public string Get()
		{
			string missionElapsedTime = String.Format("{0:00}/{1:00}:{2:00}", time.Days, time.Hours, time.Minutes);
			return (missionElapsedTime);
		}

		/// <summary>
		/// Sets the Mission Elapsed Time
		///		based on Mission Elapsed Time from the Orbit Header Column
		///		gets the day of Mission from the column before the MET Column
		/// </summary>
		/// <param name="Schedule">Array of TV Schedule read from NASA Spreadsheet</param>
		/// <param name="row"></param>
		/// <param name="column"></param>
		public void Set(Array Schedule, int row, int column)
		{
			Object cellTime = Schedule.GetValue(row, column);
			if (cellTime != null)
			{
				if (cellTime.GetType().ToString() == Properties.Resources.SYSTEM_DOUBLE)
				{
					Object cellDay = Schedule.GetValue(row, column - 1);
					if (cellDay.GetType().ToString() == Properties.Resources.SYSTEM_STRING)
					{
						time = DateTime.FromOADate((double)cellTime).TimeOfDay;
						Match mtchDay = rgxMissionDay.Match(cellDay.ToString());
						GroupCollection grpcolDay = mtchDay.Groups;
						if (grpcolDay[Properties.Resources.IX_DAY].Success)
						{
							int missionDay = Convert.ToInt32(grpcolDay[Properties.Resources.IX_DAY].ToString());
							TimeSpan tsDay = new TimeSpan(TimeSpan.TicksPerDay);
							for (int day = 1; day <= missionDay; day++)
							{
								time = time.Add(tsDay);
							}
						}
					}
				}
			}
		}
	}
	#endregion	//	MissionDuration - Handles Mission Elapsed Time

	#region NASA STS TV Schedule Entry
	/// <summary>
	/// Information returned from schedule entry
	/// </summary>
	public class NasaStsTVScheduleEntry : Object
	{
		/// <summary>
		/// Type of Schedule Entry
		/// </summary>
		ScheduleType typeEntry;
		/// <summary>
		/// Getter/Setter for type of schedule entry
		/// </summary>
		public ScheduleType TypeEntry
		{
			get { return (typeEntry); }
			private set { typeEntry = value; }
		}

		/// <summary>
		/// Begin DateTime for event
		/// </summary>
		private DateTime beginDate;
		/// <summary>
		/// Gets/Sets Beginning Date and Time for event
		/// </summary>
		public DateTime BeginDate
		{
			get { return (beginDate); }
			private set
			{
				if (value != null)
					beginDate = value;
			}
		}
		/// <summary>
		/// End DateTime for event
		/// </summary>
		private DateTime endDate;
		/// <summary>
		/// Gets/Sets Ending Date and Time for event
		/// </summary>
		public DateTime EndDate
		{
			get { return (endDate); }
			private set
			{
				if (value != null)
					endDate = value;
			}
		}
		/// <summary>
		/// Subject of the event
		/// </summary>
		private string subject;
		/// <summary>
		/// Gets/Sets the Subject for event
		/// </summary>
		public string Subject
		{
			get { return (subject); }
			private set
			{
				if (value != null)
					subject = value.Trim();
				else
					subject = null;
			}
		}
		/// <summary>
		/// Shuttle Orbit count
		/// </summary>
		private double orbit;
		/// <summary>
		/// Gets/Sets Shuttle Orbit count of the event
		/// </summary>
		public double Orbit
		{
			get { return (orbit); }
			private set { orbit = value; }
		}
		/// <summary>
		/// Site for the event
		/// </summary>
		private string site;
		/// <summary>
		/// Gets/Sets Site for the event
		/// </summary>
		public string Site
		{
			get { return (site); }
			private set
			{
				if (value != null)
					if (value.Length > 0)
						site = value.Trim();
					else
						site = "";
				else
					site = null;
			}
		}
		/// <summary>
		/// Flight Day of the event
		/// </summary>
		private string flightDay;
		/// <summary>
		/// Gets/Sets Flight Day for the event
		/// </summary>
		public string FlightDay
		{
			get { return (flightDay); }
			private set
			{
				if (value != null)
					flightDay = value.Trim();
				else
					flightDay = null;
			}
		}
		/// <summary>
		/// Was the event revised in this revision
		/// </summary>
		private bool changed;
		/// <summary>
		/// Was the event revised in the revision
		/// </summary>
		/// <returns>true if the event was revised</returns>
		public bool Revised()
		{
			return (changed);
		}
		/// <summary>
		/// Gets/Sets the revision indicator for the event
		/// </summary>
		public bool Changed
		{
			get { return (changed); }
			private set { changed = value; }
		}

		/// <summary>
		/// Constructor for the event
		/// </summary>
		/// <param name="entryBeginDateTime">Beginning Date and Time</param>
		/// <param name="entryEndDateTime">Ending Date and Time</param>
		/// <param name="entryRevised">Revision Indicator</param>
		/// <param name="entrySubject">Subject</param>
		/// <param name="entryOrbit">Orbit</param>
		/// <param name="entrySite">Site</param>
		/// <param name="entryFlightDay">Flight Day</param>
		public NasaStsTVScheduleEntry(DateTime entryBeginDateTime, DateTime entryEndDateTime, bool entryRevised,
			string entrySubject, double entryOrbit, string entrySite, string entryFlightDay, ScheduleType type)
		{
			BeginDate = entryBeginDateTime;
			EndDate = entryEndDateTime;
			Subject = entrySubject;
			Orbit = entryOrbit;
			Site = entrySite;
			FlightDay = entryFlightDay;
			Changed = entryRevised;
			TypeEntry = type;
		}
	}
	#endregion	//	NASA STS TV Schedule Entry

	#region Microsoft.Office.Interop.Excel Properties & Methods
	/// <summary>
	///
	/// </summary>
	class InteropExcelInterface : Object, IDisposable
	{
		/// <summary>
		/// InteropExcelInterface has been disposed
		/// </summary>
		bool disposed;
		/// <summary>
		/// Carriage Return, Line Feed
		/// </summary>
		private const string CRLF = "\r\n";

		/// <summary>
		/// Interop Excel Application
		/// </summary>
		private InteropExcel.ApplicationClass m_InteropExcelApplication;
		/// <summary>
		/// Gets/Sets Interop Excel Application
		/// </summary>
		private InteropExcel.ApplicationClass InteropExcelApplication
		{
			get
			{
				if (m_InteropExcelApplication == null)
				{
					m_InteropExcelApplication = new Microsoft.Office.Interop.Excel.ApplicationClass();
				}
				return (m_InteropExcelApplication);
			}
			set
			{
				if (value == null)
					m_InteropExcelApplication = null;
				else
					m_InteropExcelApplication = value;
			}
		}
		/// <summary>
		/// Interop Excel Workbook
		/// </summary>
		private InteropExcel.WorkbookClass m_InteropExcelWorkbook;
		/// <summary>
		/// Gets/Sets Interop Excel Workbook
		/// </summary>
		private InteropExcel.WorkbookClass InteropExcelWorkbook
		{
			get { return (m_InteropExcelWorkbook); }
			set
			{
				if (value == null)
					m_InteropExcelWorkbook = null;
				else
					m_InteropExcelWorkbook = value;
			}
		}
		/// <summary>
		/// Interop Excel Sheets
		/// </summary>
		private InteropExcel.Sheets m_InteropExcelSheets;
		/// <summary>
		/// Gets/Sets Interop Excel Sheets
		/// </summary>
		private InteropExcel.Sheets InteropExcelSheets
		{
			get { return (m_InteropExcelSheets); }
			set
			{
				if (value == null)
					m_InteropExcelSheets = null;
				else
					m_InteropExcelSheets = value;
			}
		}
		/// <summary>
		/// Interop Excel Worksheet
		/// </summary>
		private InteropExcel.Worksheet m_InteropExcelWorksheet;
		/// <summary>
		/// Gets/Sets Interop Excel Worksheet
		/// </summary>
		private InteropExcel.Worksheet InteropExcelWorksheet
		{
			get { return (m_InteropExcelWorksheet); }
			set
			{
				if (value == null)
					m_InteropExcelWorksheet = null;
				else
					m_InteropExcelWorksheet = value;
			}
		}
		/// <summary>
		/// Interop Excel Range
		/// </summary>
		private InteropExcel.Range m_InteropExcelRange;
		/// <summary>
		/// Gets/Sets Interop Excel Range
		/// </summary>
		private InteropExcel.Range InteropExcelRange
		{
			get { return (m_InteropExcelRange); }
			set
			{
				if (value == null)
					m_InteropExcelRange = null;
				else
					m_InteropExcelRange = value;
			}
		}

		/// <summary>
		/// Indicates if Excel has successfully opened the spreadsheet
		/// </summary>
		private bool m_SuccessfullyOpened;
		/// <summary>
		/// Gets/Sets Successful Opening of Excel Spreadsheet
		/// </summary>
		public bool SuccessfullyOpened
		{
			get { return (m_SuccessfullyOpened); }
			set { m_SuccessfullyOpened = value; }
		}

		/// <summary>
		/// Constructor for the Microsoft.Office.Interop.Excel interface
		/// </summary>
		public InteropExcelInterface()
		{
			InteropExcelRange = null;
			InteropExcelWorksheet = null;
			InteropExcelWorkbook = null;
			InteropExcelApplication = new Microsoft.Office.Interop.Excel.ApplicationClass();
		}

		/// <summary>
		///
		/// </summary>
		~InteropExcelInterface()
		{
			Dispose();
		}

		public void Close()
		{
			if (InteropExcelWorkbook != null)
				InteropExcelWorkbook.Close(false, Type.Missing, Type.Missing);
			if (InteropExcelApplication != null)
			{
				InteropExcelApplication.DisplayAlerts = false;
				InteropExcelApplication.Quit();
			}
		}

		public void Dispose()
		{
			Dispose(true);
			GC.SuppressFinalize(this);
		}

		protected virtual void Dispose(bool disposing)
		{
			if (!disposed)
			{
				if (disposing)
				{
					try
					{
						if (InteropExcelWorkbook != null)
						{
							try
							{
								//InteropExcelWorkbook.Close(false, Type.Missing, Type.Missing);
								InteropExcelWorkbook = null;
							}
							catch (COMException comExp)
							{
								if (Properties.Settings.Default.CopyExceptionsToClipboard)
									Clipboard.SetText(comExp.Message + CRLF + comExp.StackTrace, TextDataFormat.Text);
								InteropExcelWorkbook = null;
								throw;
							}
						}
						if (InteropExcelApplication != null)
						{
							try
							{
								//InteropExcelApplication.DisplayAlerts = false;
								//InteropExcelApplication.Quit();
								InteropExcelApplication = null;
							}
							catch (COMException comExp)
							{
								if (Properties.Settings.Default.CopyExceptionsToClipboard)
									Clipboard.SetText(comExp.Message + CRLF + comExp.StackTrace, TextDataFormat.Text);
								InteropExcelApplication = null;
								throw;
							}
							catch (InvalidComObjectException exp)
							{
								if (exp.Message == Properties.Resources.EXP_INVALIDCOMOBJECTEXCEPTION_RCW)
								{
									//	System.Runtime.InteropServices.InvalidComObjectException was unhandled
									//	  Message="COM object that has been separated from its underlying RCW cannot be used."
									//	  Source="Microsoft.Office.Interop.Excel"
									//	  StackTrace:
									//	       at Microsoft.Office.Interop.Excel.ApplicationClass.Quit()
									//	       at PermanentVacations.Nasa.Sts.Schedule.InteropExcelInterface.Finalize() in E:\Documents and Settings\RalphHightower\My Documents\Visual Studio 2005\Projects\NasaTvSchedule\NASA_STS_TV_Schedule\NasaStsTvSchedule.cs:line 2679
									string target = exp.TargetSite.ToString();
									if ((target != Properties.Resources.EXP_INVALIDCOMOBJECTEXCEPTION_INTEROPEXCEL_DISPLAYALERTS) &&
										(target != Properties.Resources.EXP_INVALIDCOMOBJECTEXCEPTION_INTEROPEXCEL_QUIT))
									{
										throw;
									}
								}
								else
									throw;
							}
						}
					}
					finally
					{
						InteropExcelRange = null;
						InteropExcelWorksheet = null;
						InteropExcelWorkbook = null;
						InteropExcelApplication = null;
					}
				}
				disposed = true;
			}
		}

		/// <summary>
		/// Method to open Nasa TV Schedule using Microsoft.Office.Interop.Excel
		/// </summary>
		public System.Array OpenExcelFile(string NasaTVScheduleFile)
		{
			System.Array printArea = null;

			SuccessfullyOpened = false;
			try
			{
				InteropExcelApplication = new Microsoft.Office.Interop.Excel.ApplicationClass();
				InteropExcelWorkbook = (InteropExcel.WorkbookClass)InteropExcelApplication.Workbooks.Open(NasaTVScheduleFile,
					false, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
					Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
				InteropExcelSheets = InteropExcelWorkbook.Worksheets;
				InteropExcelWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)InteropExcelSheets.get_Item(1);
				//
				//	COM Exception: Print_Area is not defined in spreadsheet (My Downloads\NASA\STS-116\tvsched_reva.xls
				//
				InteropExcelRange = InteropExcelWorksheet.get_Range(Properties.Resources.NASA_PRINT_AREA, Type.Missing);
				printArea = (System.Array)InteropExcelRange.Cells.Value2;
				//	Don't show Excel application
				InteropExcelApplication.Visible = false;
				SuccessfullyOpened = true;
				return (printArea);
			}
			catch (COMException comException)
			{
				if (comException.TargetSite.Name == Properties.Resources.EXP_COMEXCEPTION_INTEROPEXCEL_OPENEXCELFILE_GETRANGE)
				{
					string explanation = Properties.Resources.INVALIDFILEFORMAT_NO_PRINT_AREA;
					throw new InvalidFileFormatException(String.Format(explanation, NasaTVScheduleFile), comException);
				}
				else
				{
					if (Properties.Settings.Default.CopyExceptionsToClipboard)
						Clipboard.SetText(comException.Message + CRLF + comException.StackTrace, TextDataFormat.Text);
					throw;
				}
			}
		}

		/// <summary>
		/// Formats an Excel Time value (which is System.Double in .Net) to a display time
		/// </summary>
		/// <param name="row">Row of spreadsheet</param>
		/// <param name="column">Column of spreadsheet</param>
		/// <returns>Time as a string formatted similar to DateTime.ToString("hh:mm tt")</returns>
		public string FormatTime(System.Array printArea, int row, int column)
		{
			string formattedTime = "";
			if (InteropExcelApplication != null)
				formattedTime = InteropExcelApplication.WorksheetFunction.Text(printArea.GetValue(row, column), @"HH:MM AM/PM");
			return (formattedTime);
		}

	}
	#endregion	//	Microsoft.Office.Interop.Excel Properties & Methods

	#region Microsoft.Office.Tools.Excel class Properties & Methods
	class ToolsExcelInterface : Object
	{
		/// <summary>
		/// Carriage Return, Line Feed
		/// </summary>
		private const string CRLF = "\r\n";

		/// <summary>
		/// Tools Excel Workbook
		/// </summary>
		private ToolsExcel.Workbook m_ToolsExcelWorkbook;
		/// <summary>
		/// Gets/Sets Tools Excel Workbook
		/// </summary>
		private ToolsExcel.Workbook ToolsExcelWorkbook
		{
			get { return (m_ToolsExcelWorkbook); }
			set
			{
				if (value == null)
					m_ToolsExcelWorkbook = null;
				else
					m_ToolsExcelWorkbook = value;
			}
		}
		/// <summary>
		/// Tools Excel Worksheet
		/// </summary>
		private ToolsExcel.Worksheet m_ToolsExcelWorksheet;
		/// <summary>
		/// Gets/Sets Tools Excel Worksheet
		/// </summary>
		private ToolsExcel.Worksheet ToolsExcelWorksheet
		{
			get { return (m_ToolsExcelWorksheet); }
			set
			{
				if (value == null)
					m_ToolsExcelWorksheet = null;
				else
					m_ToolsExcelWorksheet = value;
			}
		}
		/// <summary>
		/// Tools Excel Range
		/// </summary>
		private ToolsExcel.NamedRange m_ToolsExceNamedlRange;
		/// <summary>
		/// Gets/Sets Tools Excel Range
		/// </summary>
		private ToolsExcel.NamedRange ToolsExcelNamedRange
		{
			get { return (m_ToolsExceNamedlRange); }
			set
			{
				if (value == null)
					m_ToolsExceNamedlRange = null;
				else
					m_ToolsExceNamedlRange = value;
			}
		}

		/// <summary>
		/// Indicates if Excel has successfully opened the spreadsheet
		/// </summary>
		private bool m_SuccessfullyOpened;
		/// <summary>
		/// Gets/Sets Successful Opening of Excel Spreadsheet
		/// </summary>
		public bool SuccessfullyOpened
		{
			get { return (m_SuccessfullyOpened); }
			set { m_SuccessfullyOpened = value; }
		}

		/// <summary>
		/// Constructor for Microsoft.Office.Tools.Excel
		/// </summary>
		public ToolsExcelInterface()
		{
			ToolsExcelWorkbook = null;
			ToolsExcelWorksheet = null;
			ToolsExcelNamedRange = null;
		}

		/// <summary>
		/// Destructor for Microsoft.Office.Tools.Excel
		/// </summary>
		~ToolsExcelInterface()
		{
			ToolsExcelWorkbook = null;
			ToolsExcelWorksheet = null;
			ToolsExcelNamedRange = null;
		}

		/// <summary>
		/// Method to open Microsoft.Office.Tools.Excel (this has problems)
		/// </summary>
		public System.Array OpenExcelFile(string NasaTVScheduleFile)
		{
			System.Array printArea = null;
			try
			{
				ToolsExcelWorkbook.OpenLinks(NasaTVScheduleFile, true, Type.Missing);
			}
			catch (COMException comExp)
			{
				if (Properties.Settings.Default.CopyExceptionsToClipboard)
					Clipboard.SetText(comExp.Message + CRLF + comExp.StackTrace);
				throw;
			}

			return (printArea);
		}

		/// <summary>
		/// Formats an Excel Time value (which is System.Double in .Net) to a display time
		/// </summary>
		/// <param name="row">Row of spreadsheet</param>
		/// <param name="weekdayMonthDay">Column of spreadsheet</param>
		/// <returns>Time as a string formatted similar to DateTime.ToString("hh:mm tt")</returns>
		public string FormatTime(System.Array printArea, int row, int cell)
		{
			string formattedTime = "";
			if (ToolsExcelWorkbook != null)
			{
				Microsoft.Office.Interop.Excel.ApplicationClass excelApplication =
					(Microsoft.Office.Interop.Excel.ApplicationClass)ToolsExcelWorkbook.Application;
				formattedTime = excelApplication.WorksheetFunction.Text(printArea.GetValue(row, cell), @"HH:MM AM/PM");
			}
			return (formattedTime);
		}
	}
	#endregion	//	Microsoft.Office.Tools.Excel class Properties & Methods
}