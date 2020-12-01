using LinqToExcel;
using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;

namespace ExcelDataLoader
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
        public readonly string Name;
        public readonly IImmutableList<(int index, int number)> Calling;
        public readonly IImmutableList<(int index, int number)> CalledBy;
        public readonly IImmutableList<(int index, int number)> CommonChanges;

        public Service(string name, IImmutableList<(int index, int number)> calling, IImmutableList<(int index, int number)> calledBy, IImmutableList<(int index, int number)> commonChanges)
        {
            Name = name;
            Calling = calling;
            CalledBy = calledBy;
            CommonChanges = commonChanges;
        }
    }

    public static class Loader
    {
        public static IImmutableList<Service> GetData()
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

            return services.Values.Select(s => new Service(
                    name: s.service.Service,
                    calling: callerCalleesServices
                        .Where(c => ReferenceEquals(c.caller.service, s.service))
                        .Select(c => (c.callee.index, number: c.numberOfCalls))
                        .ToImmutableList(),
                    calledBy: callerCalleesServices
                        .Where(c => ReferenceEquals(c.callee.service, s.service))
                        .Select(c => (c.caller.index, number: c.numberOfCalls))
                        .ToImmutableList(),
                    commonChanges: commonChangesServices
                        .Where(c => ReferenceEquals(c.service1.service, s.service) || ReferenceEquals(c.service2.service, s.service))
                        .Select(c => (index: ReferenceEquals(c.service1.service, s.service) ? c.service2.index : c.service1.index, number: c.numberOfCommonChanges))
                        .ToImmutableList()))
                .ToImmutableList();
        }
    }
}
