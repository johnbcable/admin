      <h4>Current Coaching Staff</h4>
      <table width="100%">
        <thead>
          <tr>
            <th>Name</th>
            <th>Charge rates</th>
            <th>Action</th>
          </tr>
        </thead>
        <tbody>
          {{#each items}}
          <tr>
            <td>{{forename1}} {{surname}}</td>
            <td>
              Hourly: &pound;{{hourlyrate}}<br />
              Half-hourly: &pound;{{halfhourlyrate}}
            </td>
            <td>
              <a href="/admin/#/coaches/{{uniqueref}}" class="small button coachedit" data-name="{{surname}}">Edit</a>
            </td>
          </tr>
          {{/each}}
        </tbody>
      </table>

