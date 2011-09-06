/**
 * This script synchronize events by time in Sony Vegas (H.264).
 * You can set 2 parameters (m_timeType and m_startPointType).
 * 
 * Creation date: 06 September 2011
 * Version: 1.0
 * Written by Alexander Sokol
 * If you have any questions or propositions please contact
 * with me: mailto:saneea@bk.ru
 **/

using System;
using System.Windows.Forms;
using System.IO;

using Sony.Vegas;

class EntryPoint
{
    private static TimeType m_timeType = TimeType.LastWrite; //you can use one of "TimeType.LastWrite", "TimeType.Creation" or "TimeType.LastAccess"
    private static StartPointType m_startPointType = StartPointType.Finish; //you can use one of "StartPointType.Begin" or "StartPointType.Finish"

    private enum TimeType { LastAccess, LastWrite, Creation };
    private enum StartPointType { Begin, Finish };

    private static int m_nEventCount = 0;

    public void FromVegas(Vegas vegas)
    {
        m_vegas = vegas;
        Execute();
    }

    public void Execute()
    {
        if (0 == m_vegas.Project.Tracks.Count)
        {
            MessageBox.Show("Any track was not detected");
            return;
        }

        DateTime? tempMinGlobalTime = GetEventWithMinGlobalTime(m_vegas.Project.Tracks);

        if (0 == m_nEventCount)
        {
            MessageBox.Show("Any event was not detected");
            return;
        }

        if (null == tempMinGlobalTime)
        {
            MessageBox.Show("Script can not calculate earliest time");
            return;
        }

        m_minGlobalTime = (DateTime)tempMinGlobalTime;

        foreach (TrackEventGroup group in m_vegas.Project.Groups)
        {
            HandleGroup(group);
        }

        MessageBox.Show("Script has been completed successfully");
    }

    bool HandleGroup(TrackEventGroup group)
    {
        DateTime? groupTimeNul = null;

        foreach (TrackEvent trackEvent in group)
        {
            DateTime? tempTime = GetGlobalTime(trackEvent);
            if (null != tempTime && null == groupTimeNul)
            {
                groupTimeNul = tempTime;
            }
        }

        if (null == groupTimeNul)
        {
            return false;
        }

        DateTime groupTime = (DateTime)groupTimeNul;

        TimeSpan timeSpan = new TimeSpan(groupTime.Ticks - m_minGlobalTime.Ticks);
        double dOffset = timeSpan.TotalMilliseconds;

        foreach (TrackEvent trackEvent in group)
        {
            trackEvent.Start = new Timecode(dOffset);
        }

        return true;
    }

    private static DateTime? GetEventWithMinGlobalTime(Tracks tracks)
    {
        DateTime? resultTime = null;

        foreach (Track track in tracks)
        {
            DateTime? newLeadTime = GetEventWithMinGlobalTime(track);
            if ((null == resultTime) || (null != newLeadTime && null != resultTime && newLeadTime < resultTime))
            {
                resultTime = newLeadTime;
            }
        }
        return resultTime;
    }

    private static DateTime? GetEventWithMinGlobalTime(Track track)
    {
        DateTime? resultTime = null;

        foreach (TrackEvent trackEvent in track.Events)
        {
            DateTime? newTime = GetGlobalTime(trackEvent);
            if ((null == resultTime) || (null != newTime && null != resultTime && newTime < resultTime))
            {
                resultTime = newTime;
            }

            ++m_nEventCount;
        }
        return resultTime;
    }

    private static DateTime? GetGlobalTime(TrackEvent trackEvent)
    {
        DateTime? dateTime = GetTimeFromFile(trackEvent.ActiveTake.MediaPath, m_timeType);
        if (null != dateTime && StartPointType.Finish == m_startPointType)
        {
            TimeSpan length = new TimeSpan(0, 0, 0, 0, (int)trackEvent.Length.ToMilliseconds());
            dateTime = dateTime.Value.Subtract(length);
        }
        return dateTime;
    }

    private static DateTime? GetTimeFromFile(string sFilePath, TimeType timeModel)
    {
        try
        {
            FileInfo file = new FileInfo(sFilePath);
            switch (timeModel)
            {
                case TimeType.Creation:
                    return file.CreationTime;
                case TimeType.LastAccess:
                    return file.LastAccessTime;
                case TimeType.LastWrite:
                    return file.LastWriteTime;
                default:
                    return null;
            }
        }
        catch (Exception)
        {
            return null;
        }
    }

    Vegas m_vegas;
    DateTime m_minGlobalTime;
}
