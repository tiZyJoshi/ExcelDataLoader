using LinqToExcel;
using System.Collections.Generic;
using System.Diagnostics;
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
        public List<(int index, int number)> Dependencies { get; set; }
    }

    public static class Loader
    {
        // close excel before running the program for major performance improvements!
        public static void Main()
        {
            const string fileName = @"C:\sem-dev\SWA\themonolith.xlsx";
            var excel = new ExcelQueryFactory(fileName);

            var services = excel.Worksheet<ServiceModel>("ServiceList")
                .ToList()
                .Where(s => s?.Service != null)
                .ToList();
            var serviceDictionary = services
                .Select((service, index) => (service, index))
                .ToDictionary(s => s.service.Service, s => s);

            var callercallees = excel.Worksheet<CallerCalleeModel>("CallerCallee").ToList();
            var commonChanges = excel.Worksheet<CommonChangeModel>("CommonChanges").Where(co => co.NumberOfCommonChanges > 1).ToList();

            var servicesWDeps = callercallees
                .Select(cc => serviceDictionary[cc.Caller].index)
                .Union(callercallees
                    .Select(cc => serviceDictionary[cc.Callee].index))
                .Union(commonChanges
                    .Select(cc => serviceDictionary[cc.Service1].index))
                .Union(commonChanges
                    .Select(cc => serviceDictionary[cc.Service2].index))
                .Select(i => services[i])
                .ToList();

            var serviceWDepsDictionary = servicesWDeps
                .Select((service, index) => (service, index))
                .ToDictionary(s => s.service.Service, s => s);

            var callerCalleesServices = callercallees
                .Select(c => (caller: serviceWDepsDictionary[c.Caller], callee: serviceWDepsDictionary[c.Callee], numberOfCalls: c.NumberOfCallsLastYear))
                .ToList();

            var commonChangesServices = commonChanges
                .Select(c => (service1: serviceWDepsDictionary[c.Service1], service2: serviceWDepsDictionary[c.Service2], numberOfCommonChanges: c.NumberOfCommonChanges))
                .ToList();

            var data = servicesWDeps
                .Select(s =>
                {
                    var serviceCalling = callerCalleesServices
                        .Where(c => ReferenceEquals(c.caller.service, s))
                        .Select(c => (c.callee.index, number: c.numberOfCalls))
                        .ToList();
                    var serviceCommonChanges = commonChangesServices
                        .Where(c => ReferenceEquals(c.service1.service, s))
                        .Select(c => (c.service2.index, number: c.numberOfCommonChanges))
                        .ToList();

                    var dependencies = serviceCalling
                        .Select(ca => ca.index)
                        .Union(serviceCommonChanges
                            .Select(co => co.index))
                        .Select(di =>
                            (index: di,
                            number: serviceCalling
                                        .Where(sc => sc.index == di)
                                        .Sum(sc => sc.number) +
                                    serviceCommonChanges
                                        .Where(sc => sc.index == di)
                                        .Sum(sc => sc.number)))
                        .ToList();

                    return new Service
                    {
                        Name = s.Service,
                        Dependencies = dependencies
                    };
                })
                .ToList();

            var serializer = new XmlSerializer(typeof(List<Service>));
            var writer = new StreamWriter("data.xml", false);
            serializer.Serialize(writer, data);
            writer.Close();
        }
    }
}
