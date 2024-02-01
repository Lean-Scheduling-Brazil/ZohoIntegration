using Newtonsoft.Json;
using System.Data;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.Diagnostics;
using System.Globalization;
using System.Text;


class Program
{
    static string connectionString = "server=sqlserverlsbaz.database.windows.net;database=powerbimssqlserveraz;uid=lsbsa;password=Lsb17012017;MultipleActiveResultSets=True;";

    static async Task Main()
    {
        Stopwatch stopwatch = Stopwatch.StartNew();

        using (var conn = new SqlConnection(connectionString))
        {
            await conn.OpenAsync();

            // Credenciais
            string tokenUrl = "https://accounts.zoho.com/oauth/v2/token";
            string refreshToken = "1000.da38bd0b95118893c5a9b6be10d99154.312927f32f71289347e335ddc39006aa";
            string clientId = "1000.U11OPPS6O3K37Z22ZT29O4TSVTEYYN";
            string clientSecret = "d62975aa3ec9bbdc5e039c16a29cc21cf8db70170c";
            string scope = "Desk.search.READ,Desk.tickets.READ,Desk.settings.READ,Desk.basic.READ,Desk.contacts.READ";

            // Endpoints
            string ticketsUrl = "https://desk.zoho.com/api/v1/tickets";
            string agentsUrl = "https://desk.zoho.com/api/v1/agents?from=0&limit=100";
            string contactsUrl = "https://desk.zoho.com/api/v1/contacts";
            string accountsUrl = "https://desk.zoho.com/api/v1/accounts";
            string customfielddUrl = "https://support.zoho.com/api/v1/accounts/search";

            // Obter access token
            string accessToken = await GetAccessToken(scope, tokenUrl, refreshToken, clientId, clientSecret);

            // Contagem de tickets
            string countUrl = "https://desk.zoho.com/api/v1/ticketsCount";
            int totalTickets = await GetTicketsCount(countUrl, accessToken);

            // Paginação
            int limit = 100;

            //// Truncate tables
            await Truncatetables(conn);


            var a = ProcessCustomFields(customfielddUrl, limit, accessToken, conn);
            var b = ProcessAgents(agentsUrl, accessToken, conn);
            var c = ProcessAccounts(accountsUrl, limit, accessToken, conn);
            var d = ProcessContacts(contactsUrl, limit, accessToken, conn);
            var e = ProcessTickets(ticketsUrl, totalTickets, limit, accessToken, conn);
            var f = ProcessArchivedTickets(ticketsUrl, limit, accessToken, conn);


            await Task.WhenAll(a, b, c, d, e, f);

            conn.Close();
            // Aguardar todas as tarefas serem concluídas        
            stopwatch.Stop();
            Console.WriteLine($"Tempo total de execução: {stopwatch.Elapsed}");
            Console.WriteLine("Finalizado.");

        }
    }

    // Métodos

    static async Task ProcessAgents(string agentsUrl, string accessToken, SqlConnection connection)
    {
        try
        {
            HttpResponseMessage agentsResponse = await GetData(agentsUrl, accessToken);

            if (agentsResponse.IsSuccessStatusCode)
            {
                string content = await agentsResponse.Content.ReadAsStringAsync();
                AgentsResponse data = JsonConvert.DeserializeObject<AgentsResponse>(content);

                if (data != null && data.Data.Count > 0)
                {
                    InsertAgentsDataAsync(connection, data.Data);
                }
                else
                {
                    Console.WriteLine("Nenhum agente foi encontrado ou a resposta está vazia.");
                }
            }
            else
            {
                Console.WriteLine($"Falha na requisição de agentes: {agentsResponse.StatusCode}");
            }

        }
        catch (Exception ex)
        {
            Console.WriteLine($"Erro ao processar agentes: {ex.Message}");
        }
        Console.WriteLine("Process agents finalizado");

    }
    static async Task ProcessTickets(string ticketsUrl, int totalTickets, int limit, string accessToken, SqlConnection connection)
    {
        int from = 0;
        List<Ticket> allTickets = new List<Ticket>();
        List<string> ticketIds = new List<string>(); // Lista para armazenar os IDs dos tickets

        while (from < totalTickets)
        {
            string paginatedUrl = $"{ticketsUrl}?from={from}&limit={limit}";

            try
            {
                HttpResponseMessage response = await GetData(paginatedUrl, accessToken);

                if (response.IsSuccessStatusCode)
                {
                    string content = await response.Content.ReadAsStringAsync();
                    TicketsResponse data = JsonConvert.DeserializeObject<TicketsResponse>(content);

                    if (data != null && data.Data.Count > 0)
                    {
                        allTickets.AddRange(data.Data);

                        // Adicionar os IDs dos tickets à lista
                        ticketIds.AddRange(data.Data.Select(ticket => ticket.Id));
                    }
                }
                else
                {
                    Console.WriteLine($"Erro ao buscar tickets: {response.StatusCode}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao processar tickets: {ex.Message}");
            }

            from += limit;
        }

        // Inserir todos os tickets de uma vez
        await InsertTicketsData(connection, allTickets);
        Console.WriteLine("Process tickets finalizado");


        Console.WriteLine("Process de apontamentos iniciado");
        // Processar apontamentos para todos os tickets
        await ProcessTimeEntries(connection, ticketIds, accessToken); // Passar a lista de IDs
        Console.WriteLine("Process apontamentos finalizado");


    }
    static async Task ProcessArchivedTickets(string ticketsUrl, int limit, string accessToken, SqlConnection connection)
    {
        int offset = 0;
        List<ArchivedTicket> allArchivedTickets = new List<ArchivedTicket>();
        List<string> ticketIds = new List<string>(); // Lista para armazenar os IDs dos tickets


        while (true) // Usaremos um loop infinito e controlaremos o término dentro do loop
        {
            string archivedticketsUrl = $"{ticketsUrl}/archivedTickets?departmentId=358470000000006907&from={offset}&limit={limit}";

            try
            {
                HttpResponseMessage response = await GetData(archivedticketsUrl, accessToken);

                if (response.IsSuccessStatusCode)
                {
                    string content = await response.Content.ReadAsStringAsync();
                    ArchivedTicketsResponse ticketsData = JsonConvert.DeserializeObject<ArchivedTicketsResponse>(content);

                    if (ticketsData != null && ticketsData.Data.Count > 0)
                    {
                        allArchivedTickets.AddRange(ticketsData.Data);
                        ticketIds.AddRange(ticketsData.Data.Select(ticket => ticket.Id));

                        offset += limit;
                    }
                    else
                    {
                        break; // Nenhum ticket arquivado restante, saia do loop
                    }
                }
                else
                {
                    Console.WriteLine($"Erro ao buscar tickets arquivados: {response.StatusCode}");
                    break; // Erro na solicitação, saia do loop
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao processar tickets arquivados: {ex.Message}");
                break; // Erro no processamento, saia do loop
            }
        }

        // Inserir todos os tickets arquivados de uma vez
        await InsertArchivedTicketsDataAsync(connection, allArchivedTickets);
        Console.WriteLine("Process archived tickets finalizado");


        Console.WriteLine("Process de archived apontamentos iniciado");
        // Processar apontamentos para todos os tickets
        await ProcessTimeEntries(connection, ticketIds, accessToken); // Passar a lista de IDs
        Console.WriteLine("Process archived apontamentos finalizado");


        Console.WriteLine("Processamento de tickets arquivados finalizado");
    }
    static async Task ProcessAccounts(string accountsUrl, int limit, string accessToken, SqlConnection connection)
    {
        int offset = 0;
        bool hasMoreAccounts = true;
        List<Account> allAccounts = new List<Account>();  // Lista para acumular todas as contas

        while (hasMoreAccounts)
        {
            string accountsRequestUrl = $"{accountsUrl}?include=owner&from={offset}&limit={limit}";

            try
            {
                HttpResponseMessage response = await GetData(accountsRequestUrl, accessToken);

                if (response.IsSuccessStatusCode)
                {
                    string content = await response.Content.ReadAsStringAsync();
                    AccountsResponse accountsData = JsonConvert.DeserializeObject<AccountsResponse>(content);

                    if (accountsData != null && accountsData.Data.Count > 0)
                    {
                        allAccounts.AddRange(accountsData.Data);  // Acumula os dados de contas
                        offset += limit;
                    }
                    else
                    {
                        hasMoreAccounts = false;
                    }
                }
                else
                {
                    Console.WriteLine($"Erro ao buscar contas: {response.StatusCode}");
                    hasMoreAccounts = false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao processar contas: {ex.Message}");
                hasMoreAccounts = false;
            }
        }

        if (allAccounts.Count > 0)
        {
            await InsertAccountsDataAsync(connection, allAccounts);  // Insere todas as contas de uma vez
        }

        Console.WriteLine("Process accounts finalizado");
    }
    static async Task ProcessContacts(string contactsUrl, int limit, string accessToken, SqlConnection connection)
    {
        int offset = 0;
        bool hasMoreContacts = true;
        List<Contact> allContacts = new List<Contact>();  // Lista para acumular todos os contatos

        while (hasMoreContacts)
        {
            string contactsRequestUrl = $"{contactsUrl}?from={offset}&limit={limit}";

            try
            {
                HttpResponseMessage response = await GetData(contactsRequestUrl, accessToken);

                if (response.IsSuccessStatusCode)
                {
                    string content = await response.Content.ReadAsStringAsync();
                    ContactsResponse contactsData = JsonConvert.DeserializeObject<ContactsResponse>(content);

                    if (contactsData != null && contactsData.Data.Count > 0)
                    {
                        allContacts.AddRange(contactsData.Data);  // Acumula os dados de contatos
                        offset += limit;
                    }
                    else
                    {
                        hasMoreContacts = false;
                    }
                }
                else
                {
                    Console.WriteLine($"Erro ao buscar contatos: {response.StatusCode}");
                    hasMoreContacts = false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao processar contatos: {ex.Message}");
                hasMoreContacts = false;
            }
        }

        if (allContacts.Count > 0)
        {
            await InsertContactsDataAsync(connection, allContacts);  // Insere todos os contatos de uma vez
        }

        Console.WriteLine("Process contacts finalizado");
    }
    static async Task ProcessTimeEntries(SqlConnection connection, List<string> ticketIds, string accessToken)
    {
        List<TimeEntry> allTimeEntries = new List<TimeEntry>();

        // Defina o número máximo de solicitações simultâneas
        int maxConcurrentRequests = 25;
        SemaphoreSlim semaphore = new SemaphoreSlim(maxConcurrentRequests);

        // Lista para armazenar tarefas de solicitação
        List<Task> requestTasks = new List<Task>();

        foreach (var ticketId in ticketIds)
        {
            // Aguarde o semáforo para adquirir uma ranhura antes de fazer a solicitação
            await semaphore.WaitAsync();

            // Crie uma tarefa para a solicitação
            Task requestTask = Task.Run(async () =>
            {
                Console.WriteLine($"processando apontamentos do ticket {ticketId}");

                try
                {
                    string timeEntriesUrl = $"https://desk.zoho.com/api/v1/tickets/{ticketId}/timeEntry";
                    HttpResponseMessage response = await GetData(timeEntriesUrl, accessToken);

                    if (response.IsSuccessStatusCode)
                    {
                        string content = await response.Content.ReadAsStringAsync();
                        TimeEntriesResponse timeEntriesData = JsonConvert.DeserializeObject<TimeEntriesResponse>(content);

                        if (timeEntriesData != null && timeEntriesData.Data.Count > 0)
                        {
                            lock (allTimeEntries)
                            {
                                allTimeEntries.AddRange(timeEntriesData.Data);
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine($"Erro ao buscar apontamentos do ticket {ticketId}: {response.StatusCode}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Erro ao processar apontamentos: {ex.Message}");
                }
                finally
                {
                    // Libere o semáforo após a conclusão da solicitação
                    semaphore.Release();
                }
            });

            requestTasks.Add(requestTask);
        }

        // Aguarde todas as tarefas de solicitação serem concluídas
        await Task.WhenAll(requestTasks);

        // Inserir todos os apontamentos de tempo de uma vez
        await InsertTimeEntriesDataAsync(connection, allTimeEntries);
    }
    static async Task ProcessCustomFields(string customfielddUrl, int limit, string accessToken, SqlConnection connection)
    {
        int offset = 0;
        bool hasMoreCustomFields = true;
        List<CustomFieldData> allCustomFields = new List<CustomFieldData>();

        while (hasMoreCustomFields)
        {
            string CustomFieldsRequestUrl = $"{customfielddUrl}?from={offset}&limit={limit}";

            try
            {
                HttpResponseMessage response = await GetData(CustomFieldsRequestUrl, accessToken);

                if (response.IsSuccessStatusCode)
                {
                    string content = await response.Content.ReadAsStringAsync();
                    CustomFieldsResponse customfieldsData = JsonConvert.DeserializeObject<CustomFieldsResponse>(content);

                    if (customfieldsData != null && customfieldsData.Data.Count > 0)
                    {
                        allCustomFields.AddRange(customfieldsData.Data);
                        offset += limit;
                    }
                    else
                    {
                        hasMoreCustomFields = false;
                    }
                }
                else
                {
                    Console.WriteLine($"Erro ao buscar custom fields: {response.StatusCode}");
                    hasMoreCustomFields = false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao processar custom fields: {ex.Message}");
                hasMoreCustomFields = false;
            }
        }

        if (allCustomFields.Count > 0)
        {
            await InsertCustomFieldsAsync(connection, allCustomFields);
        }

        Console.WriteLine("Process custom fields finalizado");
    }









    static async Task Truncatetables(SqlConnection connection)
    {
        using (var cmd = connection.CreateCommand())
        {
            cmd.CommandText = "TRUNCATE TABLE custom_fields";
            cmd.ExecuteNonQueryAsync();
            Console.WriteLine("custom_fields truncade");
            cmd.CommandText = "TRUNCATE TABLE tickets";
            cmd.ExecuteNonQueryAsync();
            Console.WriteLine("tickets truncade");
            cmd.CommandText = "TRUNCATE TABLE archivedtickets";
            cmd.ExecuteNonQueryAsync();
            Console.WriteLine("archivedtickets truncade");
            cmd.CommandText = "TRUNCATE TABLE agents";
            cmd.ExecuteNonQueryAsync();
            Console.WriteLine("agents truncade");
            cmd.CommandText = "TRUNCATE TABLE accounts";
            cmd.ExecuteNonQueryAsync();
            Console.WriteLine("accounts truncade");
            cmd.CommandText = "TRUNCATE TABLE contacts";
            cmd.ExecuteNonQueryAsync();
            Console.WriteLine("contacts truncade");
            cmd.CommandText = "TRUNCATE TABLE ticketstimeentries";
            cmd.ExecuteNonQueryAsync();
            Console.WriteLine("ticketstimeentries truncade");
        }
    }
    private static DateTime? ParseNullableDateTime(string dateTimeString)
    {
        if (string.IsNullOrEmpty(dateTimeString))
        {
            return null;
        }

        if (DateTime.TryParseExact(dateTimeString, "yyyy-MM-ddTHH:mm:ss.fffZ", CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out var parsedDateTime))
        {
            return parsedDateTime;
        }

        // Log or handle the case where the date string is not valid
        Console.WriteLine($"Invalid date format: {dateTimeString}");
        return null;
    }
    static async Task<HttpResponseMessage> GetData(string url, string accessToken)
    {
        var handler = new HttpClientHandler()
        {
            ServerCertificateCustomValidationCallback = HttpClientHandler.DangerousAcceptAnyServerCertificateValidator
        };
        var client = new HttpClient(handler);

        client.DefaultRequestHeaders.Add("Authorization", $"Zoho-oauthtoken {accessToken}");
        client.Timeout = TimeSpan.FromMinutes(5);

        return await client.GetAsync(url);
    }
    static async Task<string> GetAccessToken(string scope, string tokenUrl, string refreshToken, string clientId, string clientSecret)
    {
        HttpClient client = new HttpClient();

        string requestBody = $"refresh_token={refreshToken}&client_id={clientId}&client_secret={clientSecret}&grant_type=refresh_token&scope={scope}";

        HttpResponseMessage response = await client.PostAsync(tokenUrl, new StringContent(requestBody, Encoding.UTF8, "application/x-www-form-urlencoded"));

        string content = await response.Content.ReadAsStringAsync();

        return JsonHelper.GetJsonValue(content, "access_token");
    }
    static async Task<int> GetTicketsCount(string url, string accessToken)
    {
        HttpClient client = new HttpClient();
        client.DefaultRequestHeaders.Add("Authorization", $"Zoho-oauthtoken {accessToken}");

        HttpResponseMessage response = await client.GetAsync(url);

        if (response.IsSuccessStatusCode)
        {
            string json = await response.Content.ReadAsStringAsync();
            return int.Parse(JsonHelper.GetJsonValue(json, "count"));
        }

        return 0;
    }





    static async Task InsertTicketsData(SqlConnection connection, List<Ticket> tickets)
    {

        string insertQuery = @"
            INSERT INTO tickets (id, ticketNumber, email, subject, status, createdTime, priority, channel, dueDate, responseDueDate, commentCount, threadCount, closedTime, onholdTime, departmentId, contactId, assigneeId, teamId, webUrl)
            VALUES (@Id, @TicketNumber, @Email, @Subject, @Status, @CreatedTime, @Priority, @Channel, @DueDate, @ResponseDueDate, @CommentCount, @ThreadCount, @ClosedTime, @OnholdTime, @DepartmentId, @ContactId, @AssigneeId, @TeamId, @WebUrl)";

        foreach (var ticket in tickets)
        {

            if (!(await TicketsExistsAsync(connection, ticket.Id)))
            {
                try
                {
                    using (SqlCommand command = new SqlCommand(insertQuery, connection))
                    {
                        Console.WriteLine($"Writing ticket with Id: {ticket.Id}");

                        // Converta as datas para o formato correto
                        DateTime? createdTime = string.IsNullOrEmpty(ticket.CreatedTime) ? (DateTime?)null : Convert.ToDateTime(ticket.CreatedTime);
                        DateTime? dueDate = string.IsNullOrEmpty(ticket.DueDate) ? (DateTime?)null : DateTime.ParseExact(ticket.DueDate, "yyyy-MM-ddTHH:mm:ss.fffZ", CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal);
                        DateTime? responseDueDate = string.IsNullOrEmpty(ticket.ResponseDueDate) ? (DateTime?)null : DateTime.ParseExact(ticket.ResponseDueDate, "yyyy-MM-ddTHH:mm:ss.fffZ", CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal);
                        DateTime? closedTime = string.IsNullOrEmpty(ticket.ClosedTime) ? (DateTime?)null : DateTime.ParseExact(ticket.ClosedTime, "yyyy-MM-ddTHH:mm:ss.fffZ", CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal);
                        DateTime? onholdTime = string.IsNullOrEmpty(ticket.OnholdTime) ? (DateTime?)null : DateTime.ParseExact(ticket.OnholdTime, "yyyy-MM-ddTHH:mm:ss.fffZ", CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal);

                        // Defina os valores para os parâmetros
                        command.Parameters.Add("@Id", SqlDbType.BigInt).Value = Convert.ToInt64(ticket.Id);
                        command.Parameters.Add("@TicketNumber", SqlDbType.VarChar, 50).Value = ticket.TicketNumber;
                        if (ticket.Email != null)
                        {
                            command.Parameters.Add("@Email", SqlDbType.VarChar, 100).Value = ticket.Email;
                        }
                        else
                        {
                            command.Parameters.Add("@Email", SqlDbType.VarChar, 100).Value = "None";
                        }

                        command.Parameters.Add("@Subject", SqlDbType.VarChar, 200).Value = ticket.Subject;
                        command.Parameters.Add("@Status", SqlDbType.VarChar, 50).Value = ticket.Status;
                        command.Parameters.Add("@CreatedTime", SqlDbType.DateTime).Value = createdTime ?? SqlDateTime.MinValue.Value;
                        if (ticket.Priority != null)
                        {
                            command.Parameters.Add("@Priority", SqlDbType.VarChar, 50).Value = ticket.Priority;
                        }
                        else
                        {
                            command.Parameters.Add("@Priority", SqlDbType.VarChar, 50).Value = "None";
                        }
                        command.Parameters.Add("@Channel", SqlDbType.VarChar, 50).Value = ticket.Channel;
                        command.Parameters.Add("@DueDate", SqlDbType.DateTime).Value = dueDate ?? SqlDateTime.MinValue.Value;
                        command.Parameters.Add("@ResponseDueDate", SqlDbType.DateTime).Value = responseDueDate ?? SqlDateTime.MinValue.Value;
                        command.Parameters.Add("@CommentCount", SqlDbType.Int).Value = Convert.ToInt32(ticket.CommentCount);
                        command.Parameters.Add("@ThreadCount", SqlDbType.Int).Value = Convert.ToInt32(ticket.ThreadCount);
                        command.Parameters.Add("@ClosedTime", SqlDbType.DateTime).Value = closedTime ?? SqlDateTime.MinValue.Value;
                        command.Parameters.Add("@OnholdTime", SqlDbType.DateTime).Value = onholdTime ?? SqlDateTime.MinValue.Value;
                        command.Parameters.Add("@DepartmentId", SqlDbType.BigInt).Value = Convert.ToInt64(ticket.DepartmentId);
                        command.Parameters.Add("@ContactId", SqlDbType.BigInt).Value = Convert.ToInt64(ticket.ContactId);
                        command.Parameters.Add("@AssigneeId", SqlDbType.BigInt).Value = Convert.ToInt64(ticket.AssigneeId);
                        command.Parameters.Add("@TeamId", SqlDbType.BigInt).Value = Convert.ToInt64(ticket.TeamId);
                        command.Parameters.Add("@WebUrl", SqlDbType.VarChar, 500).Value = ticket.WebUrl;

                        await command.ExecuteNonQueryAsync();
                        Console.WriteLine($"Ticket with Id {ticket.Id} inserted successfully.");


                    }

                }

                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing ticket with Id {ticket.Id}: {ex.Message}");
                    // Adicione informações adicionais de depuração
                    Console.WriteLine($"Ticket data: {JsonConvert.SerializeObject(ticket)}");
                    // Você pode optar por registrar os detalhes da exceção ou manipulá-la de uma maneira que se adeque à sua aplicação
                }
            }
            else
            {
                Console.WriteLine($"Ticket with Id {ticket.Id} already exists.");
            }


        }

    }
    static async Task InsertTimeEntriesDataAsync(SqlConnection connection, List<TimeEntry> timeEntries)
    {
        string insertQuery = @"
            INSERT INTO ticketstimeentries 
            (id, ticket_number, subject, executed_time, department_id, description, owner_id, mode, is_trashed, billing_type, request_id, created_time, request_charge_type, layout_id, layout_name, seconds_spent, minutes_spent, hours_spent, is_billable, created_by, total_cost) 
            VALUES 
            (@Id, @TicketNumber, @Subject, @ExecutedTime, @DepartmentId, @Description, @OwnerId, @Mode, @IsTrashed, @BillingType, @RequestId, @CreatedTime, @RequestChargeType, @LayoutId, @LayoutName, @SecondsSpent, @MinutesSpent, @HoursSpent, @IsBillable, @CreatedBy, @TotalCost)";

        foreach (TimeEntry timeEntry in timeEntries)
        {
            using (SqlCommand command = new SqlCommand(insertQuery, connection))
            {
                command.Parameters.AddWithValue("@Id", timeEntry.Parent.Id);
                command.Parameters.AddWithValue("@TicketNumber", timeEntry.Parent.TicketNumber ?? (object)DBNull.Value);
                command.Parameters.AddWithValue("@Subject", timeEntry.Parent.Subject ?? (object)DBNull.Value);
                command.Parameters.AddWithValue("@ExecutedTime", timeEntry.ExecutedTime);
                command.Parameters.AddWithValue("@DepartmentId", timeEntry.DepartmentId);
                command.Parameters.AddWithValue("@Description", timeEntry.Description ?? (object)DBNull.Value);
                command.Parameters.AddWithValue("@OwnerId", timeEntry.OwnerId);
                command.Parameters.AddWithValue("@Mode", timeEntry.Mode ?? (object)DBNull.Value);
                command.Parameters.AddWithValue("@IsTrashed", timeEntry.IsTrashed);
                command.Parameters.AddWithValue("@BillingType", timeEntry.BillingType ?? (object)DBNull.Value);
                command.Parameters.AddWithValue("@RequestId", timeEntry.RequestId);
                command.Parameters.AddWithValue("@CreatedTime", timeEntry.CreatedTime);
                command.Parameters.AddWithValue("@RequestChargeType", timeEntry.RequestChargeType ?? (object)DBNull.Value);
                command.Parameters.AddWithValue("@LayoutId", timeEntry.LayoutId);
                command.Parameters.AddWithValue("@LayoutName", timeEntry.LayoutName ?? (object)DBNull.Value);
                command.Parameters.AddWithValue("@SecondsSpent", timeEntry.SecondsSpent);
                command.Parameters.AddWithValue("@MinutesSpent", timeEntry.MinutesSpent);
                command.Parameters.AddWithValue("@HoursSpent", timeEntry.HoursSpent);
                command.Parameters.AddWithValue("@IsBillable", timeEntry.IsBillable);
                command.Parameters.AddWithValue("@CreatedBy", timeEntry.CreatedBy);
                command.Parameters.AddWithValue("@TotalCost", timeEntry.TotalCost);

                await command.ExecuteNonQueryAsync();
            }
            Console.WriteLine($"Apontamento do ticket {timeEntry.Parent.TicketNumber} gravado com successo");
        }
    }
    static async Task InsertArchivedTicketsDataAsync(SqlConnection connection, List<ArchivedTicket> tickets)
    {
        string insertQuery = @"INSERT INTO archivedtickets (Id, TicketNumber, LayoutId, Email, Phone, Subject, Status, 
                       StatusType, CreatedTime, Category, Language, SubCategory, Priority, Channel, DueDate,
                       ResponseDueDate, CommentCount, Sentiment, ThreadCount, ClosedTime, OnholdTime, AccountId, 
                       DepartmentId, ContactId, ProductId, AssigneeId, TeamId, ChannelCode, WebUrl,
                       CustomerResponseTime)
                       VALUES (@Id, @TicketNumber, @LayoutId, @Email, @Phone, @Subject, @Status, @StatusType, 
                       @CreatedTime, @Category, @Language, @SubCategory, @Priority, @Channel, @DueDate, 
                       @ResponseDueDate, @CommentCount, @Sentiment, @ThreadCount, @ClosedTime, @OnholdTime,
                       @AccountId, @DepartmentId, @ContactId, @ProductId, @AssigneeId, @TeamId, @ChannelCode, 
                       @WebUrl, @CustomerResponseTime)";

        foreach (ArchivedTicket ticket in tickets)
        {
            try
            {
                using (SqlCommand command = new SqlCommand(insertQuery, connection))
                {
                    Console.WriteLine($"Writing archived ticket with Id: {ticket.Id}");

                    DateTime? createdTime = ParseNullableDateTime(ticket.CreatedTime);
                    DateTime? dueDate = ParseNullableDateTime(ticket.DueDate);
                    DateTime? responseDueDate = ParseNullableDateTime(ticket.ResponseDueDate);
                    DateTime? closedTime = ParseNullableDateTime(ticket.ClosedTime);
                    DateTime? onholdTime = ParseNullableDateTime(ticket.OnholdTime);
                    DateTime? customerResponseTime = ParseNullableDateTime(ticket.CustomerResponseTime);


                    command.Parameters.Add("@Id", SqlDbType.VarChar).Value = (object)ticket.Id ?? DBNull.Value;
                    command.Parameters.Add("@TicketNumber", SqlDbType.VarChar).Value = (object)ticket.TicketNumber ?? DBNull.Value;
                    command.Parameters.Add("@LayoutId", SqlDbType.VarChar).Value = (object)ticket.LayoutId ?? DBNull.Value;
                    command.Parameters.Add("@Email", SqlDbType.VarChar).Value = (object)ticket.Email ?? DBNull.Value;
                    command.Parameters.Add("@Phone", SqlDbType.VarChar).Value = (object)ticket.Phone ?? DBNull.Value;
                    command.Parameters.Add("@Subject", SqlDbType.VarChar).Value = (object)ticket.Subject ?? DBNull.Value;
                    command.Parameters.Add("@Status", SqlDbType.VarChar).Value = (object)ticket.Status ?? DBNull.Value;
                    command.Parameters.Add("@StatusType", SqlDbType.VarChar).Value = (object)ticket.StatusType ?? DBNull.Value;
                    command.Parameters.Add("@CreatedTime", SqlDbType.DateTime).Value = (object)createdTime ?? DBNull.Value;
                    command.Parameters.Add("@Category", SqlDbType.VarChar).Value = (object)ticket.Category ?? DBNull.Value;
                    command.Parameters.Add("@Language", SqlDbType.VarChar).Value = (object)ticket.Language ?? DBNull.Value;
                    command.Parameters.Add("@SubCategory", SqlDbType.VarChar).Value = (object)ticket.SubCategory ?? DBNull.Value;
                    command.Parameters.Add("@Priority", SqlDbType.VarChar).Value = (object)ticket.Priority ?? DBNull.Value;
                    command.Parameters.Add("@Channel", SqlDbType.VarChar).Value = (object)ticket.Channel ?? DBNull.Value;
                    command.Parameters.Add("@DueDate", SqlDbType.DateTime).Value = (object)dueDate ?? DBNull.Value;
                    command.Parameters.Add("@ResponseDueDate", SqlDbType.DateTime).Value = (object)responseDueDate ?? DBNull.Value;
                    command.Parameters.Add("@CommentCount", SqlDbType.VarChar).Value = (object)ticket.CommentCount ?? DBNull.Value;
                    command.Parameters.Add("@Sentiment", SqlDbType.VarChar).Value = (object)ticket.Sentiment ?? DBNull.Value;
                    command.Parameters.Add("@ThreadCount", SqlDbType.VarChar).Value = (object)ticket.ThreadCount ?? DBNull.Value;
                    command.Parameters.Add("@ClosedTime", SqlDbType.DateTime).Value = (object)closedTime ?? DBNull.Value;
                    command.Parameters.Add("@OnholdTime", SqlDbType.DateTime).Value = (object)onholdTime ?? DBNull.Value;
                    command.Parameters.Add("@AccountId", SqlDbType.VarChar).Value = (object)ticket.AccountId ?? DBNull.Value;
                    command.Parameters.Add("@DepartmentId", SqlDbType.VarChar).Value = (object)ticket.DepartmentId ?? DBNull.Value;
                    command.Parameters.Add("@ContactId", SqlDbType.VarChar).Value = (object)ticket.ContactId ?? DBNull.Value;
                    command.Parameters.Add("@ProductId", SqlDbType.VarChar).Value = (object)ticket.ProductId ?? DBNull.Value;
                    command.Parameters.Add("@AssigneeId", SqlDbType.VarChar).Value = (object)ticket.AssigneeId ?? DBNull.Value;
                    command.Parameters.Add("@TeamId", SqlDbType.VarChar).Value = (object)ticket.TeamId ?? DBNull.Value;
                    command.Parameters.Add("@ChannelCode", SqlDbType.VarChar).Value = (object)ticket.ChannelCode ?? DBNull.Value;
                    command.Parameters.Add("@WebUrl", SqlDbType.VarChar).Value = (object)ticket.WebUrl ?? DBNull.Value;
                    command.Parameters.Add("@CustomerResponseTime", SqlDbType.DateTime).Value = (object)customerResponseTime ?? DBNull.Value;


                    if (!ArchivedTicketsExistsAsync(connection, ticket.Id))
                    {
                        await command.ExecuteNonQueryAsync();
                        Console.WriteLine($"Archived Ticket with Id {ticket.Id} inserted successfully.");
                    }
                    else
                    {
                        Console.WriteLine($"Archived Ticket with Id {ticket.Id} already exists.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing Archived ticket with Id {ticket.Id}: {ex.Message}");
                // Adicione informações adicionais de depuração
                Console.WriteLine($"Archived Ticket data: {JsonConvert.SerializeObject(ticket)}");
                // Você pode optar por registrar os detalhes da exceção ou manipulá-la de uma maneira que se adeque à sua aplicação
            }
        }
    }
    public static async Task InsertCustomFieldsAsync(SqlConnection connection, List<CustomFieldData> customFieldsDataList)
    {
        try
        {

            string insertQuery = @"INSERT INTO custom_fields (id, ownerId, NomeDaEmpresa, Anotacoes, VersaoSoftware, DataRenovacao, DataAdocaoNoCS, ClassificacaoABC, Agentes, UltimoContatoGestao, DuracaoContratoVigente, NivelUtilizacao, DataAssinaturaContratoVigente, Descricao, RamoAtuacao, HorasContratadas, Dominio, TipoContrato, Produtos, ModoLicenca, VencimentoXaaS, Observacoes, NivelSatisfacao) 
                               VALUES (@id, @ownerId,@NomeDaEmpresa, @Anotacoes, @VersaoSoftware, @DataRenovacao, @DataAdocaoNoCS, @ClassificacaoABC, @Agentes, @UltimoContatoGestao, @DuracaoContratoVigente, @NivelUtilizacao, @DataAssinaturaContratoVigente, @Descricao, @RamoAtuacao, @HorasContratadas, @Dominio, @TipoContrato, @Produtos, @ModoLicenca, @VencimentoXaaS, @Observacoes, @NivelSatisfacao)";

            foreach (var customFieldData in customFieldsDataList)
            {
                var cf = customFieldData.cf;

                using (SqlCommand command = new SqlCommand(insertQuery, connection))
                {
                    command.Parameters.AddWithValue("@id", customFieldData.id ?? (object)DBNull.Value);
                    command.Parameters.AddWithValue("@ownerId", customFieldData.ownerId ?? (object)DBNull.Value);
                    command.Parameters.AddWithValue("@NomeDaEmpresa", customFieldData.accountname ?? (object)DBNull.Value);
                    command.Parameters.AddWithValue("@Anotacoes", cf.cf_anotacoes ?? (object)DBNull.Value);
                    command.Parameters.AddWithValue("@VersaoSoftware", cf.cf_versao ?? (object)DBNull.Value);
                    command.Parameters.AddWithValue("@DataRenovacao", cf.cf_renovacao ?? (object)DBNull.Value);
                    command.Parameters.AddWithValue("@DataAdocaoNoCS", cf.cf_data_de_adocao_no_cs ?? (object)DBNull.Value);
                    command.Parameters.AddWithValue("@ClassificacaoABC", cf.cf_classificacao_abc ?? (object)DBNull.Value);
                    command.Parameters.AddWithValue("@Agentes", cf.cf_agentes ?? (object)DBNull.Value);
                    command.Parameters.AddWithValue("@UltimoContatoGestao", cf.cf_ultimo_contato_gestao ?? (object)DBNull.Value);
                    command.Parameters.AddWithValue("@DuracaoContratoVigente", cf.cf_duracao_do_contrato_vigente ?? (object)DBNull.Value);
                    command.Parameters.AddWithValue("@NivelUtilizacao", cf.cf_nivel_de_utilizacao ?? (object)DBNull.Value);
                    command.Parameters.AddWithValue("@DataAssinaturaContratoVigente", cf.cf_data_de_assinatura_do_ultimo_contrato ?? (object)DBNull.Value);
                    command.Parameters.AddWithValue("@Descricao", cf.cf_descricao ?? (object)DBNull.Value);
                    command.Parameters.AddWithValue("@RamoAtuacao", cf.cf_ramo_de_atuacao ?? (object)DBNull.Value);
                    command.Parameters.AddWithValue("@HorasContratadas", cf.cf_horas_contratadas ?? (object)DBNull.Value);
                    command.Parameters.AddWithValue("@Dominio", cf.cf_dominio ?? (object)DBNull.Value);
                    command.Parameters.AddWithValue("@TipoContrato", cf.cf_tipo_de_contrato ?? (object)DBNull.Value);
                    command.Parameters.AddWithValue("@Produtos", cf.cf_produto ?? (object)DBNull.Value);
                    command.Parameters.AddWithValue("@ModoLicenca", cf.cf_modo_da_licenca ?? (object)DBNull.Value);
                    command.Parameters.AddWithValue("@VencimentoXaaS", cf.cf_vencimento_xaas ?? (object)DBNull.Value);
                    command.Parameters.AddWithValue("@Observacoes", cf.cf_observacoes ?? (object)DBNull.Value);
                    command.Parameters.AddWithValue("@NivelSatisfacao", cf.cf_nivel_de_satisfacao ?? (object)DBNull.Value);

                    await command.ExecuteNonQueryAsync();
                    //Console.WriteLine($"processing Custom Field with nome da empresa:" + customFieldData.accountname.ToString());

                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing Custom Field: {ex.Message}");
        }

    }
    static async Task InsertAgentsDataAsync(SqlConnection connection, List<Agent> agents)
    {

        string insertQuery = @"INSERT INTO agents (ID, Status, Name, Role) 
                           VALUES (@ID, @Status, @Name, @Role)";

        foreach (Agent agent in agents)
        {
            using (SqlCommand command = new SqlCommand(insertQuery, connection))
            {
                command.Parameters.Add("@ID", SqlDbType.BigInt).Value = agent.Id;
                command.Parameters.Add("@Status", SqlDbType.VarChar).Value = agent.Status;
                command.Parameters.Add("@Name", SqlDbType.VarChar).Value = agent.Name;
                if (agent.Role != null) { command.Parameters.Add("@Role", SqlDbType.VarChar).Value = agent.Role; }
                else { command.Parameters.Add("@Role", SqlDbType.VarChar).Value = "None"; }

                await command.ExecuteNonQueryAsync();
            }
        }


    }
    static async Task InsertContactsDataAsync(SqlConnection connection, List<Contact> contacts)
    {


        string insertQuery = @"INSERT INTO contacts(Id, Email, Nome, Telefone, Nome_Completo, CRM_ID, AccountId) 
                               VALUES (@Id, @Email, @Nome, @Telefone, @Nome_Completo, @CRM_ID, @AccountId)";

        foreach (Contact contact in contacts)
        {
            using (SqlCommand command = new SqlCommand(insertQuery, connection))
            {
                string fullName = contact.FirstName + " " + contact.LastName;
                string firtName = contact.FirstName ?? "None";
                string email = contact.Email ?? "None";
                string phone = contact.Phone ?? "None";
                string zohoCRMId = contact.ZohoCRMContact?.Id ?? "None";
                string accountid = contact.AccountId ?? "None";

                // Verificar se o Id já existe na tabela antes de inserir
                if (!ContactIdExists(connection, contact.Id))
                {
                    command.Parameters.Add("@Id", SqlDbType.VarChar).Value = contact.Id;
                    command.Parameters.Add("@Email", SqlDbType.VarChar).Value = email;
                    command.Parameters.Add("@Nome", SqlDbType.VarChar).Value = firtName;
                    command.Parameters.Add("@Telefone", SqlDbType.VarChar).Value = phone;
                    command.Parameters.Add("@Nome_Completo", SqlDbType.VarChar).Value = fullName;
                    command.Parameters.Add("@CRM_ID", SqlDbType.VarChar).Value = zohoCRMId;
                    command.Parameters.Add("@AccountId", SqlDbType.VarChar).Value = accountid;


                    await command.ExecuteNonQueryAsync();
                }
                // Se o Id já existe, você pode optar por atualizar o registro ou lidar de outra forma
            }
        }

    }
    static async Task InsertAccountsDataAsync(SqlConnection connection, List<Account> accounts)
    {


        string insertQuery = @"INSERT INTO accounts (Id, AccountName, Email, Website, Phone, CreatedTime, OwnerId, ZohoCRMAccountId, WebUrl, BadPercentage, OkPercentage, GoodPercentage) 
                       VALUES (@Id, @AccountName, @Email, @Website, @Phone, @CreatedTime, @OwnerId, @ZohoCRMAccountId, @WebUrl, @BadPercentage, @OkPercentage, @GoodPercentage)";

        using (SqlCommand command = new SqlCommand(insertQuery, connection))
        {
            // Adicione os parâmetros ao comando fora do loop
            command.Parameters.Add("@Id", SqlDbType.VarChar);
            command.Parameters.Add("@AccountName", SqlDbType.VarChar);
            command.Parameters.Add("@Email", SqlDbType.VarChar);
            command.Parameters.Add("@Website", SqlDbType.VarChar);
            command.Parameters.Add("@Phone", SqlDbType.VarChar);
            command.Parameters.Add("@CreatedTime", SqlDbType.DateTime);
            command.Parameters.Add("@OwnerId", SqlDbType.VarChar);
            command.Parameters.Add("@ZohoCRMAccountId", SqlDbType.VarChar);
            command.Parameters.Add("@WebUrl", SqlDbType.VarChar);
            command.Parameters.Add("@BadPercentage", SqlDbType.VarChar);
            command.Parameters.Add("@OkPercentage", SqlDbType.VarChar);
            command.Parameters.Add("@GoodPercentage", SqlDbType.VarChar);

            foreach (Account account in accounts)
            {
                // Defina os valores para os parâmetros dentro do loop
                command.Parameters["@Id"].Value = account.Id;
                command.Parameters["@AccountName"].Value = account.AccountName ?? "None";
                command.Parameters["@Email"].Value = account.Email ?? "None";
                command.Parameters["@Website"].Value = account.Website ?? "None";
                command.Parameters["@Phone"].Value = account.Phone ?? "None";
                command.Parameters["@CreatedTime"].Value = account.CreatedTime;
                command.Parameters["@OwnerId"].Value = account.OwnerId ?? "None";
                command.Parameters["@ZohoCRMAccountId"].Value = (account.ZohoCRMAccount != null) ? account.ZohoCRMAccount.Id ?? "None" : "None";
                command.Parameters["@WebUrl"].Value = account.WebUrl ?? "None";
                command.Parameters["@BadPercentage"].Value = account.CustomerHappiness.BadPercentage ?? "None";
                command.Parameters["@OkPercentage"].Value = account.CustomerHappiness.OkPercentage ?? "None";
                command.Parameters["@GoodPercentage"].Value = account.CustomerHappiness.GoodPercentage ?? "None";

                // Verificar se o Id já existe na tabela antes de inserir
                if (!AccountIdExists(connection, account.Id))
                {
                    await command.ExecuteNonQueryAsync();
                }
            }
        }

    }








    static async Task<bool> TicketsExistsAsync(SqlConnection connection, string ticketid)
    {
        string query = "SELECT COUNT(*) FROM tickets WHERE id = @Id";


        using (SqlCommand command = new SqlCommand(query, connection))
        {
            command.Parameters.Add("@Id", SqlDbType.VarChar).Value = ticketid;

            // Use ExecuteScalarAsync para operações assíncronas
            int count = (int)await command.ExecuteScalarAsync();
            return count > 0;
        }
    }
    static bool ArchivedTicketsExistsAsync(SqlConnection connection, string ticketid)
    {
        string query = "SELECT COUNT(*) FROM archivedtickets WHERE Id = @Id";

        using (SqlCommand command = new SqlCommand(query, connection))
        {
            command.Parameters.Add("@Id", SqlDbType.VarChar).Value = ticketid;
            int count = (int)command.ExecuteScalar();
            return count > 0;
        }
    }
    static bool ContactIdExists(SqlConnection connection, string contactId)
    {
        string query = "SELECT COUNT(*) FROM contacts WHERE Id = @Id";

        using (SqlCommand command = new SqlCommand(query, connection))
        {
            command.Parameters.Add("@Id", SqlDbType.VarChar).Value = contactId;
            int count = (int)command.ExecuteScalar();
            return count > 0;
        }
    }
    static bool AccountIdExists(SqlConnection connection, string accountId)
    {
        string query = "SELECT COUNT(*) FROM accounts WHERE Id = @Id";

        using (SqlCommand command = new SqlCommand(query, connection))
        {
            command.Parameters.Add("@Id", SqlDbType.VarChar).Value = accountId;
            int count = (int)command.ExecuteScalar();
            return count > 0;
        }
    }

    #region CSV
    // Metodos CSV
    static void WriteCsv(List<Ticket> tickets, string filePath, string sheetName)
    {
        using (StreamWriter writer = new StreamWriter(filePath, true, Encoding.UTF8))
        {
            foreach (Ticket ticket in tickets)
            {
                writer.WriteLine($"{ticket.Id};{ticket.TicketNumber};{ticket.Email};{ticket.Subject};{ticket.Status};{ticket.CreatedTime};{ticket.Priority};{ticket.Channel};{ticket.DueDate};{ticket.ResponseDueDate};{ticket.CommentCount};{ticket.ThreadCount};{ticket.ClosedTime};{ticket.OnholdTime};{ticket.DepartmentId};{ticket.ContactId};{ticket.AssigneeId};{ticket.TeamId};{ticket.WebUrl}");

            }
        }
    }
    static void WriteCsv(List<Agent> agents, string filePath, string sheetName)
    {
        using (StreamWriter writer = new StreamWriter(filePath, false, Encoding.UTF8))
        {

            writer.WriteLine("ID;Status;Name;Role");

            foreach (Agent agent in agents)
            {
                writer.WriteLine($"{agent.Id};{agent.Status};{agent.Name};{agent.Role}");
            }
        }
    }
    static void AppendContactsToCsv(List<Contact> contacts, string csvPath)
    {
        using (StreamWriter writer = new StreamWriter(csvPath, true))
        {
            foreach (Contact contact in contacts)
            {
                string fullName = contact.FirstName + " " + contact.LastName;
                string email = contact.Email ?? "None";
                string phone = contact.Phone ?? "None";
                string zohoCRMId = contact.ZohoCRMContact?.Id ?? "None";

                writer.WriteLine($"{contact.Id},{email},{contact.FirstName},{phone},{fullName},{zohoCRMId}");
            }
        }
    }
    static void AppendAccountsToCsv(List<Account> accounts, string csvPath)
    {
        using (StreamWriter writer = new StreamWriter(csvPath, true))
        {
            foreach (Account account in accounts)
            {
                writer.WriteLine($"{account.Id},{account.AccountName},{account.Email ?? "None"},{account.Website ?? "None"},{account.Phone ?? "None"},{account.CreatedTime},{account.OwnerId},{account.ZohoCRMAccount.Id},{account.WebUrl},{account.CustomerHappiness.BadPercentage ?? "None"},{account.CustomerHappiness.OkPercentage ?? "None"},{account.CustomerHappiness.GoodPercentage ?? "None"}");
            }
        }
    }


    #endregion

    #region Classes
    public class TimeEntriesResponse
    {
        public List<TimeEntry> Data { get; set; }
    }
    public class TimeEntry
    {
        public DateTime ExecutedTime { get; set; }
        public long DepartmentId { get; set; }
        public string Description { get; set; }
        public long OwnerId { get; set; }
        public string Mode { get; set; }
        public bool IsTrashed { get; set; }
        public string BillingType { get; set; }
        public long RequestId { get; set; }
        public DateTime CreatedTime { get; set; }
        public string RequestChargeType { get; set; }
        public long LayoutId { get; set; }
        public string LayoutName { get; set; }
        public int SecondsSpent { get; set; }
        public int MinutesSpent { get; set; }
        public int HoursSpent { get; set; }
        public bool IsBillable { get; set; }
        public long CreatedBy { get; set; }
        public decimal TotalCost { get; set; }

        // Propriedade para a relação com a tabela 'tickets'
        public Ticket Parent { get; set; }
    }
    public class Parent
    {
        public long Id { get; set; }
        public string TicketNumber { get; set; }
        public string Subject { get; set; }
    }
    public class TicketsResponse
    {
        public List<Ticket> Data { get; set; }
    }
    public class Ticket
    {
        public string Id { get; set; }
        public string TicketNumber { get; set; }
        public string LayoutId { get; set; }
        public string Email { get; set; }
        public string Phone { get; set; }
        public string Subject { get; set; }
        public string Status { get; set; }
        public string StatusType { get; set; }
        public string CreatedTime { get; set; }
        public string Category { get; set; }
        public string Language { get; set; }
        public string SubCategory { get; set; }
        public string Priority { get; set; }
        public string Channel { get; set; }
        public string DueDate { get; set; }
        public string ResponseDueDate { get; set; }
        public string CommentCount { get; set; }
        public string Sentiment { get; set; }
        public string ThreadCount { get; set; }
        public string ClosedTime { get; set; }
        public string OnholdTime { get; set; }
        public string AccountId { get; set; }
        public string DepartmentId { get; set; }
        public string ContactId { get; set; }
        public string ProductId { get; set; }
        public string AssigneeId { get; set; }
        public string TeamId { get; set; }
        public string ChannelCode { get; set; }
        public string WebUrl { get; set; }
        public bool IsSpam { get; set; }
    }
    public class ArchivedTicketsResponse
    {
        public List<ArchivedTicket> Data { get; set; }
    }
    public class ArchivedTicket
    {
        public string Id { get; set; }
        public string TicketNumber { get; set; }
        public string LayoutId { get; set; }
        public string Email { get; set; }
        public string Phone { get; set; }
        public string Subject { get; set; }
        public string Status { get; set; }
        public string StatusType { get; set; }
        public string CreatedTime { get; set; }
        public string Category { get; set; }
        public string Language { get; set; }
        public string SubCategory { get; set; }
        public string Priority { get; set; }
        public string Channel { get; set; }
        public string DueDate { get; set; }
        public string ResponseDueDate { get; set; }
        public string CommentCount { get; set; }
        public string Sentiment { get; set; }
        public string ThreadCount { get; set; }
        public string ClosedTime { get; set; }
        public string OnholdTime { get; set; }
        public string AccountId { get; set; }
        public string DepartmentId { get; set; }
        public string ContactId { get; set; }
        public string ProductId { get; set; }
        public string AssigneeId { get; set; }
        public string TeamId { get; set; }
        public string ChannelCode { get; set; }
        public string WebUrl { get; set; }
        public object LastThread { get; set; }
        public string CustomerResponseTime { get; set; }
    }
    public class AgentsResponse
    {
        public List<Agent> Data { get; set; }
    }
    public class Agent
    {
        public string Id { get; set; }
        public string Status { get; set; }
        public string Name { get; set; }
        public string Role { get; set; }
    }
    public class AccountsResponse
    {
        public List<Account> Data { get; set; }
    }
    public class Account
    {
        public string Id { get; set; }
        public string AccountName { get; set; }
        public string Email { get; set; }
        public string Website { get; set; }
        public string Phone { get; set; }
        public DateTime CreatedTime { get; set; }
        public string OwnerId { get; set; }
        public string WebUrl { get; set; }
        public ZohoCRMAccount ZohoCRMAccount { get; set; }

        public CustomerHappiness CustomerHappiness { get; set; }
    }
    public class ZohoCRMAccount
    {
        public string Id { get; set; }

    }
    public class CustomerHappiness
    {
        public string BadPercentage { get; set; }
        public string OkPercentage { get; set; }
        public string GoodPercentage { get; set; }
    }
    public class ContactsResponse
    {
        public List<Contact> Data { get; set; }
    }
    public class Contact
    {
        public string Id { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }

        public string Complete { get; set; }

        public string Email { get; set; }

        public string Phone { get; set; }

        public ZohoCRMContact ZohoCRMContact { get; set; }

        public string AccountId { get; set; }



    }
    public class ZohoCRMContact
    {
        public string Id { get; set; }

    }
    public class CustomFieldsResponse
    {
        public List<CustomFieldData> Data { get; set; }
    }
    public class CustomFieldData
    {
        public string accountname { get; set; }
        public string ownerId { get; set; }
        public string id { get; set; }
        public CFData cf { get; set; }
    }
    public class CFData
    {
        public string cf_anotacoes { get; set; }
        public string cf_nivel_de_satisfacao { get; set; }
        public string cf_modo_da_licenca { get; set; }
        public string cf_versao { get; set; }
        public string cf_ultimo_contato_gestao { get; set; }
        public string cf_duracao_do_contrato_vigente { get; set; }
        public string cf_agentes { get; set; }
        public string cf_horas_contratadas { get; set; }
        public string cf_renovacao { get; set; }
        public string cf_ramo_de_atuacao { get; set; }
        public string cf_vencimento_xaas { get; set; }
        public string cf_descricao { get; set; }
        public string cf_dominio { get; set; }
        public string cf_classificacao_abc { get; set; }
        public string cf_produto { get; set; }
        public string cf_nivel_de_utilizacao { get; set; }
        public string cf_tipo_de_contrato { get; set; }
        public string cf_data_de_assinatura_do_ultimo_contrato { get; set; }
        public string cf_nome_da_empresa { get; set; }
        public string cf_data_de_adocao_no_cs { get; set; }
        public string cf_observacoes { get; set; }
    }
    public static class JsonHelper
    {
        public static string GetJsonValue(string json, string key)
        {
            int index = json.IndexOf($"\"{key}\":", StringComparison.Ordinal);
            if (index == -1) return null;

            int startIndex = index + key.Length + 3; // Considera ": " após a chave
            int endIndex = json.IndexOfAny(new char[] { ',', '}', ']' }, startIndex);

            return json.Substring(startIndex, endIndex - startIndex).Trim(' ', '"');
        }
    }
}
#endregion
