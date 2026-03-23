using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Mvc;
using System.Web;
using System.Web.Script.Serialization;
using lamloto.Services;
using lamloto.Models;

namespace lamloto.Controllers
{
    public class MinhNgocRawController : Controller
    {
        private static readonly object NguoiLapSyncLock = new object();
        private static readonly string[] NguoiLapWeekly = new[] { "Người 1", "Người 2", "Người 3", "Người 4" };
        private static readonly Dictionary<string, string> PrizeClassMap = new Dictionary<string, string>
        {
            { "DB", "giaidb" },
            { "G1", "giai1" },
            { "G2", "giai2" },
            { "G3", "giai3" },
            { "G4", "giai4" },
            { "G5", "giai5" },
            { "G6", "giai6" },
            { "G7", "giai7" },
            { "G8", "giai8" }
        };

        public ActionResult Index()
        {
            return View();
        }

        [HttpGet]
        public JsonResult GetDaiTheoNgay(string date)
        {
            DateTime ngay;
            if (!DateTime.TryParse(date, out ngay))
            {
                return JsonLarge(new
                {
                    ok = false,
                    message = "Ngày không hợp lệ.",
                    thu = "",
                    dais = new string[0],
                    logs = new string[] { "Giá trị ngày không đúng định dạng." }
                }, JsonRequestBehavior.AllowGet);
            }

            string sourceUrl;
            List<string> logs;
            var records = FindRecordsByDate(ngay.Date, out sourceUrl, out logs);

            if (records.Count > 0)
            {
                return JsonLarge(new
                {
                    ok = true,
                    message = "Đã nạp danh sách đài theo dữ liệu Minh Ngọc.",
                    thu = GetThuText(ngay.Date),
                    dais = records.Select(x => x.Dai).Distinct().OrderBy(x => x).ToArray(),
                    sourceUrl = sourceUrl,
                    logs = logs
                }, JsonRequestBehavior.AllowGet);
            }

            var lich = GetLichMienNam(ngay.Date);
            logs.Add("Không parse được danh sách đài từ trang, đang dùng lịch cố định theo thứ.");

            return JsonLarge(new
            {
                ok = true,
                message = "Không đọc được đài trực tiếp từ trang, đã dùng lịch theo thứ.",
                thu = lich.ThuText,
                dais = lich.Dais,
                sourceUrl = sourceUrl,
                logs = logs
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public JsonResult Load(string date, string dai)
        {
            DateTime ngay;
            if (!DateTime.TryParse(date, out ngay))
            {
                return JsonLarge(new
                {
                    ok = false,
                    message = "Ngày không hợp lệ.",
                    url = "",
                    html = "",
                    text = "",
                    thu = "",
                    dais = new string[0],
                    selectedDai = dai ?? "",
                    daiHopLeTheoThu = false,
                    logs = new string[] { "Giá trị ngày không đúng định dạng." },
                    record = (object)null
                }, JsonRequestBehavior.AllowGet);
            }

            if (string.IsNullOrWhiteSpace(dai))
            {
                return JsonLarge(new
                {
                    ok = false,
                    message = "Chưa chọn đài.",
                    url = "",
                    html = "",
                    text = "",
                    thu = GetThuText(ngay.Date),
                    dais = new string[0],
                    selectedDai = "",
                    daiHopLeTheoThu = false,
                    logs = new string[] { "Anh hãy chọn đài trước khi bấm Load dữ liệu." },
                    record = (object)null
                }, JsonRequestBehavior.AllowGet);
            }

            // 1) Lấy lịch trước
            var lich = GetLichMienNam(ngay.Date);
            bool daiHopLeTheoThu = lich.Dais.Any(x =>
                string.Equals(NormalizeKey(x), NormalizeKey(dai), StringComparison.OrdinalIgnoreCase));

            // 2) Nếu sai lịch thì dừng luôn, KHÔNG gọi Minh Ngọc
            if (!daiHopLeTheoThu)
            {
                return JsonLarge(new
                {
                    ok = false,
                    message = "Đài chọn không đúng lịch của ngày này. Đã dừng, không lấy dữ liệu.",
                    url = "",
                    html = "",
                    text = "",
                    thu = lich.ThuText,
                    dais = lich.Dais,
                    selectedDai = dai,
                    daiHopLeTheoThu = false,
                    logs = new string[]
                    {
                "Đài anh chọn không nằm trong lịch mở thưởng của ngày này.",
                "Lịch hợp lệ: " + string.Join(", ", lich.Dais)
                    },
                    record = (object)null
                }, JsonRequestBehavior.AllowGet);
            }

            // 3) Chỉ khi đúng lịch mới đi lấy HTML Minh Ngọc
            string sourceUrl;
            List<string> logs;
            string html;
            var records = FindRecordsByDateWithHtml(ngay.Date, out sourceUrl, out html, out logs);

            if (records.Count == 0)
            {
                return JsonLarge(new
                {
                    ok = false,
                    message = "Đã tải HTML nhưng không parse được dữ liệu kết quả theo ngày.",
                    url = sourceUrl ?? "",
                    html = html ?? "",
                    text = HtmlToPlainText(html ?? ""),
                    thu = lich.ThuText,
                    dais = lich.Dais,
                    selectedDai = dai,
                    daiHopLeTheoThu = true,
                    logs = logs,
                    record = (object)null
                }, JsonRequestBehavior.AllowGet);
            }

            var matched = records.FirstOrDefault(x => NormalizeKey(x.Dai) == NormalizeKey(dai));
            if (matched == null)
            {
                logs.Add("Các đài parse được: " + string.Join(", ", records.Select(x => x.Dai)));
                return JsonLarge(new
                {
                    ok = false,
                    message = "Đã parse được dữ liệu ngày này nhưng không thấy đài anh chọn.",
                    url = sourceUrl ?? "",
                    html = html ?? "",
                    text = HtmlToPlainText(html ?? ""),
                    thu = lich.ThuText,
                    dais = lich.Dais,
                    selectedDai = dai,
                    daiHopLeTheoThu = true,
                    logs = logs,
                    record = (object)null
                }, JsonRequestBehavior.AllowGet);
            }

            return JsonLarge(new
            {
                ok = true,
                message = "Đã tải HTML và parse thành công kết quả cho đài đã chọn.",
                url = sourceUrl ?? "",
                html = html ?? "",
                text = HtmlToPlainText(html ?? ""),
                thu = lich.ThuText,
                dais = lich.Dais,
                selectedDai = dai,
                daiHopLeTheoThu = true,
                logs = logs,
                record = new
                {
                    Dai = matched.Dai,
                    DB = matched.DB,
                    G1 = matched.G1,
                    G2 = matched.G2,
                    G3 = matched.G3,
                    G4 = matched.G4,
                    G5 = matched.G5,
                    G6 = matched.G6,
                    G7 = matched.G7,
                    G8 = matched.G8
                }
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public JsonResult GetNguoiLapState(string date)
        {
            DateTime targetDate;
            if (!DateTime.TryParse(date, out targetDate))
            {
                targetDate = DateTime.Today;
            }

            var now = DateTime.Now;
            var currentWeekStart = StartOfWeekMonday(now);
            var targetWeekStart = StartOfWeekMonday(targetDate.Date);

            NguoiLapState state;
            lock (NguoiLapSyncLock)
            {
                state = LoadNguoiLapState();
                if (state == null)
                {
                    state = new NguoiLapState
                    {
                        BaseIndex = 0,
                        WeekStartIso = currentWeekStart.ToString("yyyy-MM-dd"),
                        Version = 1,
                        UpdatedAtIso = now.ToString("yyyy-MM-ddTHH:mm:ss")
                    };
                    SaveNguoiLapState(state);
                }

                var savedWeek = ParseIsoDate(state.WeekStartIso) ?? currentWeekStart;
                var currentBaseIdx = NormalizeNguoiLapIdx(state.BaseIndex + WeeksBetween(savedWeek, currentWeekStart));

                // tiến state tới tuần hiện tại để đồng bộ ổn định giữa client
                if (savedWeek != currentWeekStart || state.BaseIndex != currentBaseIdx)
                {
                    state.BaseIndex = currentBaseIdx;
                    state.WeekStartIso = currentWeekStart.ToString("yyyy-MM-dd");
                    state.Version = Math.Max(1, state.Version) + 1;
                    state.UpdatedAtIso = now.ToString("yyyy-MM-ddTHH:mm:ss");
                    SaveNguoiLapState(state);
                }
            }

            var currentIdx = NormalizeNguoiLapIdx(state.BaseIndex);
            var targetIdx = NormalizeNguoiLapIdx(currentIdx + WeeksBetween(currentWeekStart, targetWeekStart));

            return Json(new
            {
                ok = true,
                targetDate = targetDate.ToString("yyyy-MM-dd"),
                currentWeekStart = currentWeekStart.ToString("yyyy-MM-dd"),
                baseIndexCurrentWeek = currentIdx,
                baseNameCurrentWeek = NguoiLapWeekly[currentIdx],
                targetIndex = targetIdx,
                targetName = NguoiLapWeekly[targetIdx],
                version = state.Version,
                list = NguoiLapWeekly
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public JsonResult SetNguoiLapState(int? baseIndex, int? version)
        {
            if (!baseIndex.HasValue)
            {
                return Json(new { ok = false, message = "Thiếu baseIndex." });
            }

            var newBaseIdx = NormalizeNguoiLapIdx(baseIndex.Value);
            var now = DateTime.Now;
            var currentWeekStart = StartOfWeekMonday(now);

            lock (NguoiLapSyncLock)
            {
                var state = LoadNguoiLapState() ?? new NguoiLapState
                {
                    BaseIndex = 0,
                    WeekStartIso = currentWeekStart.ToString("yyyy-MM-dd"),
                    Version = 1,
                    UpdatedAtIso = now.ToString("yyyy-MM-ddTHH:mm:ss")
                };

                if (version.HasValue && version.Value > 0 && state.Version != version.Value)
                {
                    return Json(new
                    {
                        ok = false,
                        conflict = true,
                        message = "Dữ liệu người lập vừa được cập nhật ở nơi khác. Anh tải lại để đồng bộ."
                    });
                }

                state.BaseIndex = newBaseIdx;
                state.WeekStartIso = currentWeekStart.ToString("yyyy-MM-dd");
                state.Version = Math.Max(1, state.Version) + 1;
                state.UpdatedAtIso = now.ToString("yyyy-MM-ddTHH:mm:ss");
                SaveNguoiLapState(state);

                return Json(new
                {
                    ok = true,
                    baseIndexCurrentWeek = state.BaseIndex,
                    baseNameCurrentWeek = NguoiLapWeekly[state.BaseIndex],
                    currentWeekStart = state.WeekStartIso,
                    version = state.Version,
                    list = NguoiLapWeekly
                });
            }
        }

        private JsonResult JsonLarge(object data, JsonRequestBehavior behavior)
        {
            return new JsonResult
            {
                Data = data,
                JsonRequestBehavior = behavior,
                MaxJsonLength = int.MaxValue
            };
        }

        private DateTime StartOfWeekMonday(DateTime dt)
        {
            var d = dt.Date;
            var offset = ((int)d.DayOfWeek + 6) % 7; // Monday = 0
            return d.AddDays(-offset);
        }

        private int WeeksBetween(DateTime fromWeekStart, DateTime toWeekStart)
        {
            return (int)Math.Round((toWeekStart.Date - fromWeekStart.Date).TotalDays / 7.0);
        }

        private int NormalizeNguoiLapIdx(int idx)
        {
            var mod = idx % NguoiLapWeekly.Length;
            return mod < 0 ? mod + NguoiLapWeekly.Length : mod;
        }

        private DateTime? ParseIsoDate(string iso)
        {
            DateTime dt;
            if (DateTime.TryParseExact(iso ?? "", "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
                return dt.Date;
            return null;
        }

        private string GetNguoiLapStatePath()
        {
            return Server.MapPath("~/App_Data/nguoi_lap_state.json");
        }

        private NguoiLapState LoadNguoiLapState()
        {
            var path = GetNguoiLapStatePath();
            if (!System.IO.File.Exists(path))
                return null;

            var json = System.IO.File.ReadAllText(path);
            if (string.IsNullOrWhiteSpace(json))
                return null;

            var serializer = new JavaScriptSerializer();
            return serializer.Deserialize<NguoiLapState>(json);
        }

        private void SaveNguoiLapState(NguoiLapState state)
        {
            var path = GetNguoiLapStatePath();
            var dir = Path.GetDirectoryName(path);
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            var serializer = new JavaScriptSerializer();
            var json = serializer.Serialize(state ?? new NguoiLapState());
            System.IO.File.WriteAllText(path, json, Encoding.UTF8);
        }

        private List<ParsedRecord> FindRecordsByDate(DateTime ngay, out string sourceUrl, out List<string> logs)
        {
            string html;
            return FindRecordsByDateWithHtml(ngay, out sourceUrl, out html, out logs);
        }

        private List<ParsedRecord> FindRecordsByDateWithHtml(DateTime ngay, out string sourceUrl, out string htmlUsed, out List<string> logs)
        {
            sourceUrl = "";
            htmlUsed = "";
            logs = new List<string>();

            foreach (var url in BuildUrls(ngay))
            {
                string error;
                var html = TryDownload(url, out error);

                if (string.IsNullOrWhiteSpace(html))
                {
                    logs.Add("Không tải được " + url + ": " + error);
                    continue;
                }

                var records = ParseExactDate(html, ngay);
                if (records.Count > 0)
                {
                    sourceUrl = url;
                    htmlUsed = html;
                    logs.Add("Parse thành công từ: " + url);
                    return records;
                }

                logs.Add("Tải được HTML nhưng không parse ra dữ liệu đúng ngày từ: " + url);
            }

            return new List<ParsedRecord>();
        }

        private IEnumerable<string> BuildUrls(DateTime ngay)
        {
            var slug = ngay.ToString("dd-MM-yyyy");
            var url1 = "https://www.minhngoc.net/ket-qua-xo-so/mien-nam/" + slug + ".html";
            var url2 = "https://www.minhngoc.net/kqxs/mien-nam/" + slug + ".html";
            var live = "https://www.minhngoc.net/xo-so-truc-tiep/mien-nam.html";

            if (ngay.Date == DateTime.Today)
                return new[] { live, url1, url2 };

            return new[] { url1, url2, live };
        }

        private string TryDownload(string url, out string error)
        {
            error = "";
            try
            {
                var req = (HttpWebRequest)WebRequest.Create(url);
                req.Method = "GET";
                req.Timeout = 25000;
                req.UserAgent = "Mozilla/5.0";
                req.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";
                req.Headers.Add("Accept-Language", "vi,vi-VN;q=0.9,en;q=0.8");
                req.Referer = "https://www.minhngoc.net/";

                using (var resp = (HttpWebResponse)req.GetResponse())
                using (var stream = resp.GetResponseStream())
                using (var reader = new StreamReader(stream, Encoding.UTF8, true))
                {
                    return reader.ReadToEnd();
                }
            }
            catch (WebException ex)
            {
                error = ex.Message;
                return null;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return null;
            }
        }

        private List<ParsedRecord> ParseExactDate(string html, DateTime targetDate)
        {
            var doc = new HtmlDocument();
            doc.LoadHtml(html);

            var box = PickExactDateBox(doc, targetDate);
            var scope = box ?? doc.DocumentNode;

            var tables = scope.Descendants("table")
                .Where(TableHasPrizeNumbers)
                .ToList();

            if (!tables.Any())
                return new List<ParsedRecord>();

            var records = new List<ParsedRecord>();

            foreach (var tb in tables)
            {
                var dai = GetDaiName(tb);
                if (string.IsNullOrWhiteSpace(dai))
                    continue;

                var rec = new ParsedRecord
                {
                    Dai = dai,
                    DB = GetPrizeValue(tb, "DB"),
                    G1 = GetPrizeValue(tb, "G1"),
                    G2 = GetPrizeValue(tb, "G2"),
                    G3 = GetPrizeValue(tb, "G3"),
                    G4 = GetPrizeValue(tb, "G4"),
                    G5 = GetPrizeValue(tb, "G5"),
                    G6 = GetPrizeValue(tb, "G6"),
                    G7 = GetPrizeValue(tb, "G7"),
                    G8 = GetPrizeValue(tb, "G8")
                };

                if (HasAnyPrize(rec))
                    records.Add(rec);
            }

            return records
                .GroupBy(x => NormalizeKey(x.Dai))
                .Select(g => g.First())
                .ToList();
        }

        private HtmlNode PickExactDateBox(HtmlDocument doc, DateTime targetDate)
        {
            var dateText = targetDate.ToString("dd/MM/yyyy");
            var dateSlug = targetDate.ToString("dd-MM-yyyy");

            var boxes = doc.DocumentNode.Descendants("div")
                .Where(x => HasClass(x, "box_kqxs"))
                .ToList();

            foreach (var box in boxes)
            {
                var titleNode = box.Descendants().FirstOrDefault(x => HasClass(x, "title"));
                if (titleNode == null)
                    continue;

                foreach (var a in titleNode.Descendants("a"))
                {
                    var txt = CleanText(a.InnerText);
                    var href = a.GetAttributeValue("href", "");

                    if (string.Equals(txt, dateText, StringComparison.OrdinalIgnoreCase) ||
                        (!string.IsNullOrWhiteSpace(href) && href.Contains(dateSlug)))
                    {
                        return box;
                    }
                }
            }

            if (boxes.Count == 1)
                return boxes[0];

            return null;
        }

        private bool TableHasPrizeNumbers(HtmlNode table)
        {
            foreach (var cls in PrizeClassMap.Values)
            {
                var td = table.Descendants("td").FirstOrDefault(x => HasClass(x, cls));
                if (td != null && !string.IsNullOrWhiteSpace(JoinNumbersFromTd(td)))
                    return true;
            }
            return false;
        }

        private string GetDaiName(HtmlNode table)
        {
            var tinhTd = table.Descendants("td").FirstOrDefault(x => HasClass(x, "tinh"));
            if (tinhTd == null)
                return null;

            var a = tinhTd.Descendants("a").FirstOrDefault();
            return CleanText(a != null ? a.InnerText : tinhTd.InnerText);
        }

        private string GetPrizeValue(HtmlNode table, string code)
        {
            string prizeClass;
            if (!PrizeClassMap.TryGetValue(code, out prizeClass))
                return "";

            var td = table.Descendants("td").FirstOrDefault(x => HasClass(x, prizeClass));
            return td == null ? "" : JoinNumbersFromTd(td);
        }

        private string JoinNumbersFromTd(HtmlNode td)
        {
            if (td == null) return "";

            var results = new List<string>();

            // 1) Ưu tiên tách theo "khối hiển thị" trong HTML: div, span, br, p...
            var html = td.InnerHtml ?? "";

            // biến các điểm xuống dòng/đóng khối thành dấu phân cách
            html = Regex.Replace(html, @"<br\s*/?>", "|", RegexOptions.IgnoreCase);
            html = Regex.Replace(html, @"</(div|p|li|span)>", "|", RegexOptions.IgnoreCase);

            // bỏ các tag còn lại
            html = Regex.Replace(html, @"<[^>]+>", " ");
            html = HtmlEntity.DeEntitize(html);

            var parts = html
                .Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(x => CleanText(x))
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .ToList();

            foreach (var part in parts)
            {
                var digits = DigitsOnly(part);
                if (!string.IsNullOrWhiteSpace(digits))
                    results.Add(digits);
            }

            // 2) Nếu đã tách được nhiều cụm thì trả về luôn
            if (results.Count > 1)
                return string.Join(",", results);

            // 3) Fallback: thử lấy theo các node con có text số
            results.Clear();
            var childBlocks = td.ChildNodes
                .Where(x => x.NodeType == HtmlNodeType.Element || x.NodeType == HtmlNodeType.Text)
                .Select(x => CleanText(x.InnerText))
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .ToList();

            foreach (var block in childBlocks)
            {
                var digits = DigitsOnly(block);
                if (!string.IsNullOrWhiteSpace(digits))
                    results.Add(digits);
            }

            if (results.Count > 1)
                return string.Join(",", results);

            // 4) Cuối cùng mới gom regex toàn ô
            var text = CleanText(td.InnerText);
            var matches = Regex.Matches(text, @"\d+");
            var nums = new List<string>();

            foreach (Match m in matches)
                nums.Add(m.Value);

            return string.Join(",", nums);
        }

        private string HtmlToPlainText(string html)
        {
            if (string.IsNullOrWhiteSpace(html))
                return "";

            var s = html;
            s = Regex.Replace(s, "<script[\\s\\S]*?</script>", "", RegexOptions.IgnoreCase);
            s = Regex.Replace(s, "<style[\\s\\S]*?</style>", "", RegexOptions.IgnoreCase);
            s = Regex.Replace(s, "<[^>]+>", " ");
            s = HtmlEntity.DeEntitize(s);
            s = Regex.Replace(s, "\\s+", " ").Trim();
            return s;
        }

        private string CleanText(string input)
        {
            return HtmlEntity.DeEntitize(input ?? "")
                .Replace("\u00A0", " ")
                .Replace("\r", " ")
                .Replace("\n", " ")
                .Trim();
        }

        private string DigitsOnly(string input)
        {
            return Regex.Replace(input ?? "", @"[^\d]", "");
        }

        private bool HasClass(HtmlNode node, string className)
        {
            var cls = node.GetAttributeValue("class", "");
            if (string.IsNullOrWhiteSpace(cls))
                return false;

            var parts = cls.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            return parts.Any(x => string.Equals(x, className, StringComparison.OrdinalIgnoreCase));
        }

        private bool HasAnyPrize(ParsedRecord rec)
        {
            return !string.IsNullOrWhiteSpace(rec.DB)
                || !string.IsNullOrWhiteSpace(rec.G1)
                || !string.IsNullOrWhiteSpace(rec.G2)
                || !string.IsNullOrWhiteSpace(rec.G3)
                || !string.IsNullOrWhiteSpace(rec.G4)
                || !string.IsNullOrWhiteSpace(rec.G5)
                || !string.IsNullOrWhiteSpace(rec.G6)
                || !string.IsNullOrWhiteSpace(rec.G7)
                || !string.IsNullOrWhiteSpace(rec.G8);
        }

        private string NormalizeKey(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return "";

            var s = input.Trim().ToLowerInvariant().Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder();

            foreach (var ch in s)
            {
                var uc = CharUnicodeInfo.GetUnicodeCategory(ch);
                if (uc != UnicodeCategory.NonSpacingMark)
                    sb.Append(ch);
            }

            return Regex.Replace(sb.ToString(), @"[^a-z0-9]+", "");
        }

        private string GetThuText(DateTime ngay)
        {
            switch (ngay.DayOfWeek)
            {
                case DayOfWeek.Monday: return "Thứ hai";
                case DayOfWeek.Tuesday: return "Thứ ba";
                case DayOfWeek.Wednesday: return "Thứ tư";
                case DayOfWeek.Thursday: return "Thứ năm";
                case DayOfWeek.Friday: return "Thứ sáu";
                case DayOfWeek.Saturday: return "Thứ bảy";
                default: return "Chủ nhật";
            }
        }

        private LichMienNam GetLichMienNam(DateTime ngay)
        {
            switch (ngay.DayOfWeek)
            {
                case DayOfWeek.Monday:
                    return new LichMienNam("Thứ hai", new[] { "TP. HCM", "Đồng Tháp", "Cà Mau" });
                case DayOfWeek.Tuesday:
                    return new LichMienNam("Thứ ba", new[] { "Bến Tre", "Vũng Tàu", "Bạc Liêu" });
                case DayOfWeek.Wednesday:
                    return new LichMienNam("Thứ tư", new[] { "Đồng Nai", "Cần Thơ", "Sóc Trăng" });
                case DayOfWeek.Thursday:
                    return new LichMienNam("Thứ năm", new[] { "Tây Ninh", "An Giang", "Bình Thuận" });
                case DayOfWeek.Friday:
                    return new LichMienNam("Thứ sáu", new[] { "Vĩnh Long", "Bình Dương", "Trà Vinh" });
                case DayOfWeek.Saturday:
                    return new LichMienNam("Thứ bảy", new[] { "TP. HCM", "Long An", "Bình Phước", "Hậu Giang" });
                default:
                    return new LichMienNam("Chủ nhật", new[] { "Tiền Giang", "Kiên Giang", "Đà Lạt" });
            }
        }

        private class LichMienNam
        {
            public string ThuText { get; set; }
            public string[] Dais { get; set; }

            public LichMienNam(string thuText, string[] dais)
            {
                ThuText = thuText;
                Dais = dais ?? new string[0];
            }
        }


        // [HttpPost]
        //public JsonResult UploadLiclDocx(HttpPostedFileBase liclFile)
        //{
        //    try
        //    {
        //        if (liclFile == null || liclFile.ContentLength <= 0)
        //        {
        //            return Json(new
        //            {
        //                ok = false,
        //                message = "Chưa chọn file DOCX."
        //            });
        //        }

        //        var ext = Path.GetExtension(liclFile.FileName);
        //        if (!string.Equals(ext, ".docx", StringComparison.OrdinalIgnoreCase))
        //        {
        //            return Json(new
        //            {
        //                ok = false,
        //                message = "Chỉ chấp nhận file .docx"
        //            });
        //        }

        //        var uploadDir = Server.MapPath("~/App_Data/Uploads");
        //        if (!Directory.Exists(uploadDir))
        //            Directory.CreateDirectory(uploadDir);

        //        var fileName = "LICL_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".docx";
        //        var fullPath = Path.Combine(uploadDir, fileName);

        //        liclFile.SaveAs(fullPath);

        //        Session["LICL_DOCX_PATH"] = fullPath;
        //        Session["LICL_DOCX_NAME"] = liclFile.FileName;

        //        return Json(new
        //        {
        //            ok = true,
        //            message = "Tải file thành công.",
        //            fileName = liclFile.FileName,
        //            savedAs = fileName
        //        });
        //    }
        //    catch (Exception ex)
        //    {
        //        return Json(new
        //        {
        //            ok = false,
        //            message = "Lỗi upload file: " + ex.Message
        //        });
        //    }
        //}












        [HttpPost]
        public JsonResult UploadLiclDocx(HttpPostedFileBase liclFile)
        {
            try
            {
                if (liclFile == null || liclFile.ContentLength <= 0)
                {
                    return Json(new { ok = false, message = "Chưa chọn file DOCX." });
                }

                var ext = Path.GetExtension(liclFile.FileName);
                if (!string.Equals(ext, ".docx", StringComparison.OrdinalIgnoreCase))
                {
                    return Json(new { ok = false, message = "Chỉ chấp nhận file .docx" });
                }

                var uploadDir = GetLiclUploadDir();
                if (!Directory.Exists(uploadDir))
                    Directory.CreateDirectory(uploadDir);

                var fileName = "LICL_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".docx";
                var fullPath = Path.Combine(uploadDir, fileName);

                liclFile.SaveAs(fullPath);

                Session["LICL_DOCX_PATH"] = fullPath;
                Session["LICL_DOCX_NAME"] = liclFile.FileName;

                return Json(new
                {
                    ok = true,
                    message = "Tải file LICL thành công.",
                    fileName = liclFile.FileName,
                    savedAs = fileName,
                    uploadDir = uploadDir,
                    fullPath = fullPath
                });
            }
            catch (Exception ex)
            {
                return Json(new
                {
                    ok = false,
                    message = "Lỗi upload file LICL: " + ex.Message
                });
            }
        }




        //[HttpGet]
        //public JsonResult GetLiclByDate(string date)
        //{
        //    DateTime ngay;
        //    if (!DateTime.TryParse(date, out ngay))
        //    {
        //        return Json(new
        //        {
        //            ok = false,
        //            message = "Ngày không hợp lệ.",
        //            hoiDong = new string[0],
        //            banLanhDao = new string[0],
        //            fileName = ""
        //        }, JsonRequestBehavior.AllowGet);
        //    }

        //    var path = Session["LICL_DOCX_PATH"] as string;
        //    var fileName = Session["LICL_DOCX_NAME"] as string;

        //    if (string.IsNullOrWhiteSpace(path) || !System.IO.File.Exists(path))
        //    {
        //        return Json(new
        //        {
        //            ok = false,
        //            message = "Chưa có file LICL.docx được upload.",
        //            hoiDong = new string[0],
        //            banLanhDao = new string[0],
        //            fileName = fileName ?? ""
        //        }, JsonRequestBehavior.AllowGet);
        //    }

        //    try
        //    {
        //        var service = new LiclDocxService();
        //        var info = service.ReadByDate(path, ngay.Date);

        //        return Json(new
        //        {
        //            ok = true,
        //            message = "Đọc file LICL thành công.",
        //            hoiDong = info.HoiDongNames.ToArray(),
        //            banLanhDao = info.BanLanhDaoNames.ToArray(),
        //            fileName = fileName ?? ""
        //        }, JsonRequestBehavior.AllowGet);
        //    }
        //    catch (Exception ex)
        //    {
        //        return Json(new
        //        {
        //            ok = false,
        //            message = "Lỗi đọc file LICL: " + ex.Message,
        //            hoiDong = new string[0],
        //            banLanhDao = new string[0],
        //            fileName = fileName ?? ""
        //        }, JsonRequestBehavior.AllowGet);
        //    }
        //}









        [HttpGet]
        public JsonResult GetLiclByDate(string date)
        {
            DateTime ngay;
            if (!DateTime.TryParse(date, out ngay))
            {
                return Json(new
                {
                    ok = false,
                    message = "Ngày không hợp lệ.",
                    hoiDong = new string[0],
                    banLanhDao = new string[0],
                    fileName = ""
                }, JsonRequestBehavior.AllowGet);
            }

            var path = Session["LICL_DOCX_PATH"] as string;
            var fileName = Session["LICL_DOCX_NAME"] as string;

            if (string.IsNullOrWhiteSpace(path) || !System.IO.File.Exists(path))
            {
                var dir = GetLiclUploadDir();
                path = GetLatestFilePathFromDir(dir, "LICL_*.docx");
                fileName = GetLatestFileNameFromDir(dir, "LICL_*.docx");
            }

            if (string.IsNullOrWhiteSpace(path) || !System.IO.File.Exists(path))
            {
                path = GetLatestLiclDocxPath();
                fileName = GetLatestLiclDocxName();
            }

            if (string.IsNullOrWhiteSpace(path) || !System.IO.File.Exists(path))
            {
                return Json(new
                {
                    ok = false,
                    message = "Chưa có file LICL.docx đã lưu.",
                    hoiDong = new string[0],
                    banLanhDao = new string[0],
                    fileName = ""
                }, JsonRequestBehavior.AllowGet);
            }

            try
            {
                var service = new lamloto.Services.LiclDocxService();
                var info = service.ReadByDate(path, ngay.Date);

                return Json(new
                {
                    ok = true,
                    message = (info.HoiDongNames != null && info.HoiDongNames.Count > 0)
                        ? "Đọc file LICL thành công."
                        : "Không tìm thấy Hội đồng cho ngày " + ngay.ToString("dd/MM/yyyy"),
                    hoiDong = (info.HoiDongNames ?? new System.Collections.Generic.List<string>()).ToArray(),
                    banLanhDao = (info.BanLanhDaoNames ?? new System.Collections.Generic.List<string>()).ToArray(),
                    fileName = fileName ?? ""
                }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new
                {
                    ok = false,
                    message = "Lỗi đọc file LICL: " + ex.Message,
                    hoiDong = new string[0],
                    banLanhDao = new string[0],
                    fileName = fileName ?? ""
                }, JsonRequestBehavior.AllowGet);
            }
        }









        private string GetLatestLiclDocxPath()
        {
            var uploadDir = Server.MapPath("~/App_Data/Uploads");
            if (!Directory.Exists(uploadDir))
                return null;

            var file = new DirectoryInfo(uploadDir)
                .GetFiles("*.docx")
                .OrderByDescending(f => f.LastWriteTime)
                .FirstOrDefault();

            return file != null ? file.FullName : null;
        }

        private string GetLatestLiclDocxName()
        {
            var uploadDir = Server.MapPath("~/App_Data/Uploads");
            if (!Directory.Exists(uploadDir))
                return null;

            var file = new DirectoryInfo(uploadDir)
                .GetFiles("*.docx")
                .OrderByDescending(f => f.LastWriteTime)
                .FirstOrDefault();

            return file != null ? file.Name : null;
        }



























        [HttpPost]
        public JsonResult UploadBldDocx(HttpPostedFileBase bldFile)
        {
            try
            {
                if (bldFile == null || bldFile.ContentLength <= 0)
                {
                    return Json(new { ok = false, message = "Chưa chọn file BLD .docx" });
                }

                var ext = Path.GetExtension(bldFile.FileName);
                if (!string.Equals(ext, ".docx", StringComparison.OrdinalIgnoreCase))
                {
                    return Json(new { ok = false, message = "File Ban lãnh đạo phải là .docx" });
                }

                var uploadDir = GetBldUploadDir();
                if (!Directory.Exists(uploadDir))
                    Directory.CreateDirectory(uploadDir);

                var fileName = "BLD_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".docx";
                var fullPath = Path.Combine(uploadDir, fileName);

                bldFile.SaveAs(fullPath);

                Session["BLD_DOCX_PATH"] = fullPath;
                Session["BLD_DOCX_NAME"] = bldFile.FileName;

                return Json(new
                {
                    ok = true,
                    message = "Tải file BLD thành công.",
                    fileName = bldFile.FileName,
                    savedAs = fileName,
                    uploadDir = uploadDir,
                    fullPath = fullPath
                });
            }
            catch (Exception ex)
            {
                return Json(new
                {
                    ok = false,
                    message = "Lỗi upload file BLD: " + ex.Message
                });
            }
        }
        [HttpGet]
        public JsonResult GetBldByDate(string date)
        {
            DateTime ngay;
            if (!DateTime.TryParse(date, out ngay))
            {
                return Json(new
                {
                    ok = false,
                    message = "Ngày không hợp lệ.",
                    tenLanhDao = "",
                    rawText = "",
                    fileName = ""
                }, JsonRequestBehavior.AllowGet);
            }

            var path = Session["BLD_DOCX_PATH"] as string;
            var fileName = Session["BLD_DOCX_NAME"] as string;

            if (string.IsNullOrWhiteSpace(path) || !System.IO.File.Exists(path))
            {
                var dir = GetBldUploadDir();
                path = GetLatestFilePathFromDir(dir, "BLD_*.docx");
                fileName = GetLatestFileNameFromDir(dir, "BLD_*.docx");
            }

            if (string.IsNullOrWhiteSpace(path) || !System.IO.File.Exists(path))
            {
                return Json(new
                {
                    ok = false,
                    message = "Chưa có file BLD .docx đã lưu.",
                    tenLanhDao = "",
                    rawText = "",
                    fileName = ""
                }, JsonRequestBehavior.AllowGet);
            }

            try
            {
                var service = new BldDocxService();
                var info = service.ReadByDate(path, ngay.Date);

                return Json(new
                {
                    ok = true,
                    message = string.IsNullOrWhiteSpace(info.TenLanhDao)
                        ? "Không tìm thấy Ban lãnh đạo cho ngày " + ngay.ToString("dd/MM/yyyy")
                        : "Đọc file BLD thành công.",
                    tenLanhDao = info.TenLanhDao ?? "",
                    rawText = info.RawText ?? "",
                    fileName = fileName ?? ""
                }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new
                {
                    ok = false,
                    message = "Lỗi đọc file BLD: " + ex.Message,
                    tenLanhDao = "",
                    rawText = "",
                    fileName = fileName ?? ""
                }, JsonRequestBehavior.AllowGet);
            }
        }

        private string GetLatestUploadedFilePath(string pattern)
        {
            var uploadDir = Server.MapPath("~/App_Data/Uploads");
            if (!Directory.Exists(uploadDir))
                return null;

            var file = new DirectoryInfo(uploadDir)
                .GetFiles(pattern)
                .OrderByDescending(f => f.LastWriteTime)
                .FirstOrDefault();

            return file != null ? file.FullName : null;
        }

        private string GetLatestUploadedFileName(string pattern)
        {
            var uploadDir = Server.MapPath("~/App_Data/Uploads");
            if (!Directory.Exists(uploadDir))
                return null;

            var file = new DirectoryInfo(uploadDir)
                .GetFiles(pattern)
                .OrderByDescending(f => f.LastWriteTime)
                .FirstOrDefault();

            return file != null ? file.Name : null;
        }


































        private string GetLiclUploadDir()
        {
            return Server.MapPath("~/App_Data/Uploads/LICL");
        }

        private string GetBldUploadDir()
        {
            return Server.MapPath("~/App_Data/Uploads/BLD");
        }

        private string GetLatestFilePathFromDir(string dirPath, string pattern)
        {
            if (!Directory.Exists(dirPath))
                return null;

            var file = new DirectoryInfo(dirPath)
                .GetFiles(pattern)
                .OrderByDescending(f => f.LastWriteTime)
                .FirstOrDefault();

            return file != null ? file.FullName : null;
        }

        private string GetLatestFileNameFromDir(string dirPath, string pattern)
        {
            if (!Directory.Exists(dirPath))
                return null;

            var file = new DirectoryInfo(dirPath)
                .GetFiles(pattern)
                .OrderByDescending(f => f.LastWriteTime)
                .FirstOrDefault();

            return file != null ? file.Name : null;
        }















        private class ParsedRecord
        {
            public string Dai { get; set; }
            public string DB { get; set; }
            public string G1 { get; set; }
            public string G2 { get; set; }
            public string G3 { get; set; }
            public string G4 { get; set; }
            public string G5 { get; set; }
            public string G6 { get; set; }
            public string G7 { get; set; }
            public string G8 { get; set; }
        }

        private class NguoiLapState
        {
            public int BaseIndex { get; set; }
            public string WeekStartIso { get; set; }
            public int Version { get; set; }
            public string UpdatedAtIso { get; set; }
        }
    }
}
