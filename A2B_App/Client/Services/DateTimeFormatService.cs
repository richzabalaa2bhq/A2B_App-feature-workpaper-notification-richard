using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace A2B_App.Client.Services
{
    public class DateTimeFormatService
    {
        public bool IsBetween(TimeSpan time, TimeSpan startTime, TimeSpan endTime)
        {
            if (time == startTime) return true;
            if (time == endTime) return true;

            if (startTime <= endTime)
                return (time >= startTime && time <= endTime);
            else
                return !(time >= endTime && time <= startTime);
        }
    }
}
