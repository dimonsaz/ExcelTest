using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Net.WebSockets;
using System.Threading;
using System.Diagnostics;
using System.Web.Script.Serialization;
using ExchangeSharp;

namespace ExcelTest
{

    /// <summary>
    /// 
    /// TODO:
    ///  1. support heartbeat GDX
    ///  2. test all possible disconnects GDX
    ///  3. make VBA object "Contract" and test it
    ///  4. add support for other  exchanges
    /// </summary>

    /*
    At first I just want to use the market data.  It should be analogous to connecting to a TT gateway.
Market data doesn't require authentication.

should be able to return something like the following
Contract.GetPrice(BookSide.Bid) returns object containing Best Bid Price and Size  (similar for BookSide.Ask)
Contract.GetDepth(BookSide.Bid) returns List of Price / Size objects

Contract object can be created using Exchange and Symbol.

The data should be made accessible through Excel using RTD calls that would look like this:

=GetPrice("Bithumb", "XRP", "Bid")
=GetSize("Bithumb", "XRP", "Bid")

The adapters should handle recovery from bad data states (BitMEX is good with this).  Use the websocket interface whenever possible.  REST when that's the only option.
I can give more details later, but they shouldn't impact the initial design.
Yes, set up a private Git repo, we can move it to ours later.


Priority:
GDAX (Coinbase)
BitMEX --
BitStamp
Bithumb
Bitfinex
BitFlyer --
Bittrex
Poloniex

kraken
gemini

  */
    class GdaxTickerMsg
    {
        public string type { get; set; }
        public ulong sequence { get; set; }

        public string product_id { get; set; }

        public string price { get; set; }
        public string size { get; set; }

        public string time { get; set; }

        public string side { get; set; }
        public string best_bid { get; set; }
        public string best_ask { get; set; }
    }

    public interface IConnector
    {
        void Connect();
        NumberProvider Subscribe(string instument, string side);
       // void Receive();
    }
    public class GdaxConnector : IConnector
    {
        ClientWebSocket socket;
        Task connectTask;
      //  private IRTDUpdateEvent excelCallback;
        static Dictionary<string, Market> subscribed = new Dictionary<string, Market>();

        public void Connect()
        {
            if (socket == null)
            {
                socket = new ClientWebSocket();
                connectTask = socket.ConnectAsync(new Uri("wss://ws-feed.gdax.com"), CancellationToken.None);
               
                Debug.Print("GDX socket " + socket.State);
                //Receive();
                if (readThread == null)
                {
                    readThread = new Thread(new ThreadStart(ReceiverThread));
                    readThread.Start();
                }
            }
        }

        
        public NumberProvider Subscribe(string instument, string side)
        {
            if( !subscribed.TryGetValue(instument, out Market  market) )
                subscribed[instument] = new Market();

            SubscribeAsync(instument);

            return side.StartsWith("B",StringComparison.InvariantCultureIgnoreCase) ? subscribed[instument].Bid :
                side.StartsWith("A", StringComparison.InvariantCultureIgnoreCase) ? subscribed[instument].Ask : NumberProvider.NanNumberProvider;
        }

        public async Task SubscribeAsync(string instument)
        {
           // Debug.Print("Subscribe to {0} {1} {2}", topicId, exch, instr, side);
     //       Debug.Print("Subscribe to {0}", instument);
            connectTask.Wait();
            string json = @"{""product_ids"":[""" + instument + @"""],""type"":""subscribe"",""channels"": [""ticker""]}";
            byte[] bytes = Encoding.UTF8.GetBytes(json);
            ArraySegment<byte> subscriptionMessageBuffer = new ArraySegment<byte>(bytes);
            await socket.SendAsync(subscriptionMessageBuffer, WebSocketMessageType.Text, true, CancellationToken.None);
       //     Debug.Print("SendAsync completed " + json);
        }

        Thread readThread;
        //void Receive()
        //{
        //    if (readThread == null)
        //    {
               
        //        readThread = new Thread(new ThreadStart(ReceiverThread));
        //        readThread.Start();
        //    }
        //}
        void ReceiverThread()
        {
            try
            {
                connectTask.Wait();
                while (true)
                {
                    byte[] recBytes = new byte[3 * 1024];
                    var t = new ArraySegment<byte>(recBytes);

                    var receiveAsync = socket.ReceiveAsync(t, CancellationToken.None);
                    var s = receiveAsync.Status;

                    receiveAsync.Wait();
                    var ddds = receiveAsync.Status;
                    string jsonString;
                    jsonString = Encoding.UTF8.GetString(recBytes, 0, receiveAsync.Result.Count);


                   //  Debug.Print(">> {0}", jsonString);

                    var msg = new JavaScriptSerializer().Deserialize<GdaxTickerMsg>(jsonString);

              //      Debug.Print(">>>>> {0} / {1}\t{2}", msg.best_bid, msg.best_ask, msg.product_id);

                    if (msg.product_id != null && subscribed.TryGetValue(msg.product_id, out Market m))
                    {
                        double bid, ask;
                        double.TryParse(msg.best_bid, out bid);
                        double.TryParse(msg.best_ask, out ask);
                        m.Bid.price = bid;
                        m.Ask.price = ask;
                    }

                }

            }
            catch (Exception ex)
            {
                Debug.Print(ex.ToString());
            }
        }
    }

    public class ConnectorBase
    {
        System.Timers.Timer timer = new System.Timers.Timer();
        public ConnectorBase()
        {
            timer.Interval = 15000;
            timer.Elapsed += Timer_Elapsed;
            timer.Start();
        }

        List<ConnectorBase> connectors = new List<ConnectorBase>();
        protected void Register(ConnectorBase con)
        {
            if( !connectors.Contains(con))
                connectors.Add(con);
        }
        private bool inTimer = false;
        private void Timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            try
            {
                if (inTimer)
                    return;
                inTimer = true;
                Debug.Print("in timer");
                foreach (var con in connectors.ToArray())
                {
                    con.GetTickers();
                }
            }
            finally
            {
                inTimer = false;
            }
        }
        public void Connect()
        {
        }

        protected ExchangeAPI api;
        protected decimal div = 1;
        Dictionary<string, Market> subscribed = new Dictionary<string, Market>();
        public NumberProvider Subscribe(string instument, string side)
        {
            lock (subscribed)
            {
                if (!subscribed.TryGetValue(instument, out Market market))
                    subscribed[instument] = new Market();
            }
            return side.StartsWith("B", StringComparison.InvariantCultureIgnoreCase) ? subscribed[instument].Bid :
               side.StartsWith("A", StringComparison.InvariantCultureIgnoreCase) ? subscribed[instument].Ask : NumberProvider.NanNumberProvider;

        }

        void GetTickers()
        {
            //lock (subscribed)
            {
                Debug.Print("getTickers: " + api.Name);
                foreach (var item in subscribed.ToArray<KeyValuePair<string,Market>>())
                {
                    try
                    {
                        Debug.Write("\t " + item.Key);
                        //Task<ExchangeTicker> ticker = api.GetTickerAsync(item.Key);
                        //ticker.Wait();
                        var ticker = api.GetTicker(item.Key);
                        item.Value.Bid.price = (double)(ticker.Bid / div);
                        item.Value.Ask.price = (double)(ticker.Ask / div);
                        Debug.Print(" " + item.Value.Bid.price + " / " + item.Value.Ask.price);
                    }
                    catch(Exception ex)
                    {
                        Debug.Print(item.Key + " ---> " + ex.Message);
                    }
                }
            }
        }
    }

    public class GdaxConnector2 : ConnectorBase, IConnector
    {
        public GdaxConnector2()
        {
            api = new ExchangeGdaxAPI();
            var symbols = api.GetSymbols();
            Debug.Print("{0}: ", api.Name);
            foreach (var s in symbols)
            {
                Debug.Print("\t{0}", s);
            }
            base.Register(this);
        }
    }

    public class BitStampConnector : ConnectorBase, IConnector
    {
        public BitStampConnector()
        {
            api = new ExchangeBitstampAPI();
            var symbols = api.GetSymbols();
            Debug.Print("{0}: ", api.Name);
            foreach (var s in symbols)
            {
         //       Debug.Print("\t{0}", s);
            }
            base.Register(this);
        }
    }
    public class BithumbConnector : ConnectorBase, IConnector
    {
        public BithumbConnector()
        {
            api = new ExchangeBithumbAPI();
            var symbols = api.GetSymbols();
            Debug.Print("{0}: ", api.Name);
            foreach (var s in symbols)
            {
            //    Debug.Print("\t{0}", s);
            }
            div = 1000;
            base.Register(this);
        }
    }
    public class BitfinexConnector : ConnectorBase, IConnector
    {
        public BitfinexConnector()
        {
            api = new ExchangeBitfinexAPI();
            var symbols = api.GetSymbols();
            Debug.Print("{0}: ", api.Name);
            foreach (var s in symbols)
            {
         //       Debug.Print("\t{0}",s);
            }
            base.Register(this);
        }
    }
    public class BittrexConnector : ConnectorBase, IConnector
    {
        public BittrexConnector()
        {
            api = new ExchangeBittrexAPI();
            var symbols = api.GetSymbols();
            Debug.Print("{0}: ", api.Name);
            foreach (var s in symbols)
            {
           //     Debug.Print("\t{0}", s);
            }
            base.Register(this);
        }
    }
    public class BinanceConnector : ConnectorBase, IConnector
    {
        public BinanceConnector()
        {
            api = new ExchangeBinanceAPI();
            var symbols = api.GetSymbols();
            Debug.Print("{0}: ", api.Name);
            foreach (var s in symbols)
            {
          //      Debug.Print("\t{0}", s);
            }
            base.Register(this);
        }
    }
    public class PoloniexConnector : ConnectorBase, IConnector
    {
        public PoloniexConnector()
        {
            api = new ExchangePoloniexAPI();
            var symbols = api.GetSymbols();
            Debug.Print("{0}: ", api.Name);
            foreach (var s in symbols)
            {
                Debug.Print("\t{0}", s);
            }
            //  base.Register(this);
        }
    }
    public class KrakenConnector : ConnectorBase, IConnector
    {
        public KrakenConnector()
        {
            api = new ExchangeKrakenAPI();
            var symbols = api.GetSymbols();
            Debug.Print("{0}: ", api.Name);
            foreach (var s in symbols)
            {
                Debug.Print("\t{0}", s);
            }
            base.Register(this);
        }
    }
    public class GeminiConnector : ConnectorBase, IConnector
    {
        public GeminiConnector()
        {
            api = new ExchangeGeminiAPI();
            var symbols = api.GetSymbols();
            Debug.Print("{0}: ", api.Name);
            foreach (var s in symbols)
            {
                Debug.Print("\t{0}", s);
            }
            base.Register(this);
        }
    }

    public static class ExchConnectorFactory
    {
      //  public static IRTDUpdateEvent excelCallback;
        static Dictionary<string, IConnector> exchanges = new Dictionary<string, IConnector>();
        public static IConnector Get(string exch)
        {
            exch = exch.ToUpper();
            if (exchanges.TryGetValue(exch, out IConnector con))
                return con;
            switch (exch)
            {
                case "GDX":
                   // con = new GdaxConnector();
                    con = new GdaxConnector2();
                    break;
                case "BITSTAMP":
                    con = new BitStampConnector();
                    break;
                case "BITHUMP":
                    con = new BithumbConnector();
                    break;
                case "BINANCE":
                    con = new BinanceConnector();
                    break;
                case "BITFINEX":
                    con = new BitfinexConnector();
                    break;
                case "BITREX":
                    con = new BittrexConnector();
                    break;
                case "POLONIEX":
                    con = new PoloniexConnector();
                    break;
                case "KRAKEN":
                    con = new KrakenConnector();
                    break;
                case "GEMINI":
                    con = new GeminiConnector();
                    break;


                // = RTD("cryptoprice.get",, "BitStamp", "BTC-USD", "Bid")
                default:
                    throw new Exception("Exchange " + exch + " is not yet supported");
            }
            exchanges[exch] = con;
            con.Connect();
            return con;
        }
    }

    public class Market
    {
        public Market()
        {
            Bid = new NumberProvider();
            Ask = new NumberProvider();
        }
        public NumberProvider Bid { get; set; }
        public NumberProvider Ask { get; set; }
        public override string ToString()
        {
            return Bid + " / " + Ask;
        }
    }
    public class NumberProvider
    {
        private static NumberProvider nan = new NumberProvider() { price = double.NaN };
        public static NumberProvider NanNumberProvider { get { return nan; } }
        public double price = double.NaN;
        public object Get()
        {
            return price;
        }
        public override string ToString()
        {
            return price.ToString();
        }
    }

    enum TickerSide { Bid, Ask, Last }
    public class ExtTicker
    {
        public string Symbol { get; set; }
        public IExchangeAPI API { get; set; }
        public DateTime LastUpdated { get; set; }
        public ExchangeTicker Ticker { get; set; } = new ExchangeTicker();
        public ExtTicker GetTicker()
        {
            this.Ticker = API.GetTicker(Symbol);
            LastUpdated = DateTime.Now;
      //      Debug.Print($"getticker {API.Name} {Symbol} {Ticker} ");

            Task.Run(() => this.GetTicker());
            return this;
        }
    }

    public class TickerProcessor
    {
        private Dictionary<string, ExtTicker> tickers = new Dictionary<string, ExtTicker>();
        private Dictionary<int, Tuple<TickerSide, ExtTicker>> topics = new Dictionary<int, Tuple<TickerSide, ExtTicker>>();// topicID to IsBuy? and ticker

        private string GetKey(string exch, string instument)
        {
            return exch + "+" + instument;
        }
        private TickerSide GetSide(string side)
        {
            if (side.StartsWith("B", StringComparison.InvariantCultureIgnoreCase))
                return TickerSide.Bid;
            else if(side.StartsWith("A", StringComparison.InvariantCultureIgnoreCase))
                return TickerSide.Ask;
            else if (side.StartsWith("L", StringComparison.InvariantCultureIgnoreCase))
                return TickerSide.Last;
            else
                throw new ArgumentException("side must be Bid/Ask/Last, but you used " + side, "side");
        }

        internal void Subscribe(string exch, string instument, string side, int topicId)
        {
            if(!tickers.TryGetValue(GetKey(exch,instument), out ExtTicker extTicker ))
            {
                extTicker = new ExtTicker()
                {
                    API = ExchangeAPI.GetExchangeAPI(exch),
                    Symbol = instument,
                };
                tickers[GetKey(exch, instument)] = extTicker;
            }

            if (topics.ContainsKey(topicId))
                throw new ArgumentException("topicid " + topicId + " already been used");
            topics[topicId] = new Tuple<TickerSide, ExtTicker>(GetSide(side), extTicker);
            Task.Run(() => extTicker.GetTicker());
        }

        static Dictionary<string, decimal> miltipliers = new Dictionary<string, decimal>()
        {
            { ExchangeName.Bithumb, 0.001m }
        };
        public decimal GetTopic(int topicId)
        {
            if (topics.ContainsKey(topicId))
            {
                var pair = topics[topicId];
                decimal m = 1;
                if( miltipliers.TryGetValue(pair.Item2.API.Name, out decimal mm) )
                {
                    m = mm;
                }

                switch(pair.Item1)
                {
                    case TickerSide.Bid:
                        return pair.Item2.Ticker.Bid * m;
                    case TickerSide.Ask:
                        return pair.Item2.Ticker.Ask * m;
                    case TickerSide.Last:
                        return pair.Item2.Ticker.Last * m;
                    default:
                        return decimal.Zero;
                }
                 // pair.Item2.Ticker.Volume.PriceAmount
                 // pair.Item2.Ticker.Volume.PriceSymbol
                 // pair.Item2.Ticker.Volume.QuantityAmount
                 // pair.Item2.Ticker.Volume.QuantitySymbol
            }
            else
                return decimal.Zero;
        }

        internal void Unsubscribe(int topicID)
        {
            Debug.Print($"unsubscribe from topic {topicID}");
            if (topics.TryGetValue(topicID, out Tuple<TickerSide, ExtTicker> value))
            {
                bool onlyOneTopicRemainsForThisTicker = true;
                foreach (var item in topics)
                {
                    if (item.Value.Item2 == value.Item2 && item.Value.Item1 != value.Item1)
                    {
                        onlyOneTopicRemainsForThisTicker = false;
                        break;
                    }
                }
                if (onlyOneTopicRemainsForThisTicker)
                    tickers.Remove(GetKey(value.Item2.API.Name, value.Item2.Symbol));
            }
            topics.Remove(topicID);
        }
    }

    [
       Guid  ("D146024E-64F7-4C76-AB27-D165C9610316"),
       ProgId("CryptoPrice.Get"),
    ]
    public class RTDFunctions : IRtdServer
    {
        TickerProcessor processor = new TickerProcessor();
        private IRTDUpdateEvent callback;
        private System.Windows.Forms.Timer timer;
        // Dictionary<int, NumberProvider> topics = new Dictionary<int, NumberProvider>();
        List<int> topics = new List<int>();
       
        public int ServerStart(IRTDUpdateEvent callback)
        {
            this.callback = callback;
            timer = new System.Windows.Forms.Timer();
            timer.Tick += new EventHandler(TimerEventHandler);

            timer.Interval = 1000;
            return 1;
        }

        private void TimerEventHandler(object sender, EventArgs args)
        {
            timer.Stop();
            callback.UpdateNotify();
        }

        public dynamic ConnectData(int topicId, ref Array strings, ref bool GetNewValues)
        {
            try
            {
                // Debug.Print("ConnectData: topicId " + topicId );
                string exch = strings.GetValue(0).ToString();
                string instument = strings.GetValue(1).ToString();
                string side = strings.GetValue(2).ToString();

                //var connector = ExchConnectorFactory.Get(exch);
                //var np = connector.Subscribe(instument,side);
                //topics[topicId] = np;
                //m_timer.Start();

                //  tst.Run();

                processor.Subscribe(exch, instument, side, topicId);
                // topics[topicId] = null;
                topics.Add(topicId);
                timer.Start();

                return "connecting";
            }
            catch(Exception ex)
            {
                return ex.Message;
            }
        }

        public Array RefreshData(ref int topicCount)
        {
            object[,] data = new object[2, topics.Count];
            int index = 0;
            foreach (int topicId in topics)
            {
                data[0, index] = topicId;
                // data[1, index] = topics[topicId].Get();
                data[1, index] = processor.GetTopic(topicId);
                ++index;
            }
            topicCount = topics.Count;

            timer.Start();
            return data;
        }

        public void DisconnectData(int topicID)
        {
            topics.Remove(topicID);
            processor.Unsubscribe(topicID);
        }

        public int Heartbeat()
        {
            return 1;
        }

        public void ServerTerminate()
        {
            if (null != timer)
            {
                timer.Dispose();
                timer = null;
            }
        }
    }

    /*
       public class RTDFunctions : IRtdServer
    {
        private IRTDUpdateEvent m_callback;
        private System.Windows.Forms.Timer m_timer;
        private int m_topicId;


        public int ServerStart(IRTDUpdateEvent callback)
        {
            m_callback = callback;
            m_timer = new System.Windows.Forms.Timer();
            m_timer.Tick += new EventHandler(TimerEventHandler);
            //m_timer.Elapsed += M_timer_Elapsed;
            m_timer.Interval = 2000;
            return 1;
        }

        private void TimerEventHandler(object sender, EventArgs args)
        {
            m_timer.Stop();
            m_callback.UpdateNotify();
        }

        public dynamic ConnectData(int topicId, ref Array Strings, ref bool GetNewValues)
        {
            m_topicId = topicId;
            m_timer.Start();
            return GetTime();
        }

        public Array RefreshData(ref int topicCount)
        {
            object[,] data = new object[2, 1];
            data[0, 0] = m_topicId;
            data[1, 0] = GetTime();

            topicCount = 1;

            m_timer.Start();
            return data;
        }
        private string GetTime()
        {
            return DateTime.Now.ToString("GDX!!!!: hh:mm:ss.fff");
        }

        public void DisconnectData(int TopicID)
        {
            m_timer.Stop();
        }

        public int Heartbeat()
        {
            return 1;
        }

        public void ServerTerminate()
        {
            if (null != m_timer)
            {
                m_timer.Dispose();
                m_timer = null;
            }
        }
    }
     
     */
}
