using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MySeleniumProject
{
    public class MarketData
    {
        public string Date { get; set; }
        public string Symbol { get; set; }
        public double Open { get; set; }
        public double High { get; set; }
        public double Low { get; set; }
        public double Close { get; set; }
        public long Volume { get; set; }
        public long Open_interest { get; set; }
        public double Delta { get; set; }
        public double MaxDelta { get; set; }
        public double MinDelta { get; set; }
        public double CumDelta { get; set; }
        public long BuyVolume { get; set; }
        public long SellVolume { get; set; }
        public double Vwap { get; set; }
        public double BuyVwap { get; set; }
        public double SellVwap { get; set; }
        // Cumulative fields
        public double CumVolume { get; set; }
        public double CumBuyVolume { get; set; }
        public double CumSellVolume { get; set; }
        public double CumOpen { get; set; }
        public double CumHigh { get; set; }
        public double CumLow { get; set; }
        public double CumClose { get; set; }
    }
}
