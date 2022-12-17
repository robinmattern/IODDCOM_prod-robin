pJSON =
{
  projects: [
    {
      Id: 1,
      Name: 'NATO Summit',
      Client: 'Departments of State and Defense',
      ClientWeb: 'http://www.state.;v',
      ProjectWeb: null,
      Location: 'Washington, DC',
      ProjectType: 'Internet - Multi-user LAN - Windows NT - Access 97',
      Industry: ';vt Summit',
      Description: "The project supported the NATO Summit's 50th Anniversary meeting. The system provided credentials for 2,000 foreign dignitaries and 10,000 US citizens. The Summit staff was able to monitor Airport arrivals and departures , Hotel and Event information. The Operations Center was managed using a real-time Incident System from their browser. The technology used was NT 4.0 and Active Server Pages.",
      CreatedAt: '2020-11-30 00:00:00',
      UpdatedAt: '2020-11-30 00:00:00'
    },
    {
      Id: 2,
      Name: 'Web Time Cards',
      Client: 'KPMG',
      ClientWeb: 'http://www.kpmg.com/',
      ProjectWeb: null,
      Location: 'Arlington, VA',
      ProjectType: 'Internet Data Processing - Windows 95/NT - Access ',
      Industry: 'Time and Billing',
      Description: 'The TIMEX system manages cards from KPMG employees and contractors. Individuals input their time sheet data over the Internet through a secure SSL connection. The data is analyzed and reported via an Access 97 application that is also connected via the Internet.',
      CreatedAt: '2020-11-30 00:00:00',
      UpdatedAt: '2020-11-30 00:00:00'
    }
  ]
}

  if (process) { console.log( require('util').inspect( pJSON, { depth: 99 } ) ) }