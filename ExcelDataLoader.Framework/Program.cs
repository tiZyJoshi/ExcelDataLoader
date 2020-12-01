using LinqToExcel;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;

namespace ExcelDataLoader.FrameWork
{
    class ServiceModel
    {
        public string Service { get; set; }
    }

    class CallerCalleeModel
    {
        public string Caller { get; set; }
        public string Callee { get; set; }
        public int NumberOfCallsLastYear { get; set; }
    }

    class CommonChangeModel
    {
        public string Service1 { get; set; }
        public string Service2 { get; set; }
        public int NumberOfCommonChanges { get; set; }
    }

    public class Service
    {
        public string Name { get; set; }
        public List<(int index, int number)> Calling { get; set; }
        public List<(int index, int number)> CalledBy { get; set; }
        public List<(int index, int number)> CommonChanges { get; set; }
    }

    public static class Loader
    {
        public static void Main()
        {
            var fileName = @"C:\sem-dev\SWA\themonolith.xlsx";
            var excel = new ExcelQueryFactory(fileName);
            var services = excel.Worksheet<ServiceModel>("ServiceList")
                .ToList()
                .Where(s => s?.Service != null)
                .Select((service, index) => (service, index))
                .ToDictionary(s => s.service.Service, s => s);
            var callercallees = excel.Worksheet<CallerCalleeModel>("CallerCallee").ToList();
            var commonChanges = excel.Worksheet<CommonChangeModel>("CommonChanges").ToList();

            var callerCalleesServices = callercallees
                .Select(c => (caller: services[c.Caller], callee: services[c.Callee], numberOfCalls: c.NumberOfCallsLastYear))
                .ToList();
            var commonChangesServices = commonChanges
                .Where(c => c.NumberOfCommonChanges > 2)
                .Select(c => (service1: services[c.Service1], service2: services[c.Service2], numberOfCommonChanges: c.NumberOfCommonChanges))
                .ToList();

            var data = services.Values.Select(s => new Service {
                    Name = s.service.Service,
                    Calling = callerCalleesServices
                        .Where(c => ReferenceEquals(c.caller.service, s.service))
                        .Select(c => (c.callee.index, number: c.numberOfCalls))
                        .ToList(),
                    CalledBy = callerCalleesServices
                        .Where(c => ReferenceEquals(c.callee.service, s.service))
                        .Select(c => (c.caller.index, number: c.numberOfCalls))
                        .ToList(),
                    CommonChanges = commonChangesServices
                        .Where(c => ReferenceEquals(c.service1.service, s.service) || ReferenceEquals(c.service2.service, s.service))
                        .Select(c => (index: ReferenceEquals(c.service1.service, s.service) ? c.service2.index : c.service1.index, number: c.numberOfCommonChanges))
                        .ToList()})
                .ToList();

            var serializer = new XmlSerializer(typeof(List<Service>));
            var writer = new StreamWriter("data.xml");
            serializer.Serialize(writer, data);
            writer.Close();
        }
    }
}
