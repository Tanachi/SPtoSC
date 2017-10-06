# SPtoSC
Imports data from sharepoint to sharpcloud

### How to install from Visual Studio

Create a new C# console application with the name SPtoSC.

In the project folder, replace program.cs and app.config with the ones from this repo.

#Add References

Project -> Add References

System.Configuration

#Install Packages

Tools -> Nuget Package Manager -> Package Manager Console 

Enter these lines in the console in this order.

Install-Package Microsoft.Sharepoint.2013.Client.16

Install-Package Microsoft.Sharepoint

Install-Package SharpCloud.ClientAPI

Enter your Sharpcloud and sharepoint info and the sharepoint site and list name in the app.config file.
