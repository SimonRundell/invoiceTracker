#nullable enable

namespace RenewalBilling.Forms;

partial class ReviewForm
{

    private Panel _topPanel = null!;
    private FlowLayoutPanel _topFlowPanel = null!;
    private Label _runDateLabel = null!;
    private Label _rpiLabel = null!;
    private NumericUpDown _rpiNumericUpDown = null!;
    private Label _rpiPeriodLabel = null!;
    private Label _rpiSourceLabel = null!;
    private Label _warningBanner = null!;

    private SplitContainer _mainSplitContainer = null!;

    private DataGridView _candidatesGrid = null!;
    private DataGridViewCheckBoxColumn _approveColumn = null!;
    private DataGridViewTextBoxColumn _clientColumn = null!;
    private DataGridViewTextBoxColumn _dueDateColumn = null!;
    private DataGridViewTextBoxColumn _statusColumn = null!;
    private DataGridViewTextBoxColumn _netColumn = null!;
    private DataGridViewTextBoxColumn _totalColumn = null!;

    private Panel _detailPanel = null!;
    private GroupBox _addressGroupBox = null!;
    private TableLayoutPanel _addressLayout = null!;
    private Label _clientNameLabel = null!;
    private Label _contactNameLabel = null!;
    private Label _address1Label = null!;
    private Label _address2Label = null!;
    private Label _townPostcodeLabel = null!;
    private Label _emailLabel = null!;

    private DataGridView _itemsGrid = null!;
    private DataGridViewTextBoxColumn _descriptionColumn = null!;
    private DataGridViewTextBoxColumn _qtyColumn = null!;
    private DataGridViewTextBoxColumn _currentPriceColumn = null!;
    private DataGridViewTextBoxColumn _proposedPriceColumn = null!;
    private DataGridViewTextBoxColumn _upliftColumn = null!;

    private TableLayoutPanel _totalsLayout = null!;
    private Label _subtotalLabel = null!;
    private Label _vatLabel = null!;
    private Label _totalLabel = null!;

    private Panel _bottomPanel = null!;
    private Button _createButton = null!;
    private Button _closeButton = null!;
    private Button _openFolderButton = null!;
    private ProgressBar _buildProgressBar = null!;
    private Label _statusLabel = null!;
    private TextBox _summaryTextBox = null!;

    private void InitializeComponent()
    {
        _topPanel = new Panel();
        _topFlowPanel = new FlowLayoutPanel();
        _runDateLabel = new Label();
        _rpiLabel = new Label();
        _rpiNumericUpDown = new NumericUpDown();
        _rpiPeriodLabel = new Label();
        _rpiSourceLabel = new Label();
        _warningBanner = new Label();

        _mainSplitContainer = new SplitContainer();

        _candidatesGrid = new DataGridView();
        _approveColumn = new DataGridViewCheckBoxColumn();
        _clientColumn = new DataGridViewTextBoxColumn();
        _dueDateColumn = new DataGridViewTextBoxColumn();
        _statusColumn = new DataGridViewTextBoxColumn();
        _netColumn = new DataGridViewTextBoxColumn();
        _totalColumn = new DataGridViewTextBoxColumn();

        _detailPanel = new Panel();
        _addressGroupBox = new GroupBox();
        _addressLayout = new TableLayoutPanel();
        _clientNameLabel = new Label();
        _contactNameLabel = new Label();
        _address1Label = new Label();
        _address2Label = new Label();
        _townPostcodeLabel = new Label();
        _emailLabel = new Label();

        _itemsGrid = new DataGridView();
        _descriptionColumn = new DataGridViewTextBoxColumn();
        _qtyColumn = new DataGridViewTextBoxColumn();
        _currentPriceColumn = new DataGridViewTextBoxColumn();
        _proposedPriceColumn = new DataGridViewTextBoxColumn();
        _upliftColumn = new DataGridViewTextBoxColumn();

        _totalsLayout = new TableLayoutPanel();
        _subtotalLabel = new Label();
        _vatLabel = new Label();
        _totalLabel = new Label();

        _bottomPanel = new Panel();
        _createButton = new Button();
        _closeButton = new Button();
        _openFolderButton = new Button();
        _buildProgressBar = new ProgressBar();
        _statusLabel = new Label();
        _summaryTextBox = new TextBox();

        // Top bar
        _topPanel.Dock = DockStyle.Top;
        _topPanel.Height = 76;
        _topPanel.Controls.Add(_topFlowPanel);
        _topPanel.Controls.Add(_warningBanner);

        _topFlowPanel.Dock = DockStyle.Top;
        _topFlowPanel.Height = 40;
        _topFlowPanel.FlowDirection = FlowDirection.LeftToRight;
        _topFlowPanel.WrapContents = false;
        _topFlowPanel.Padding = new Padding(8, 10, 8, 4);
        _topFlowPanel.Controls.Add(_runDateLabel);
        _topFlowPanel.Controls.Add(_rpiLabel);
        _topFlowPanel.Controls.Add(_rpiNumericUpDown);
        _topFlowPanel.Controls.Add(_rpiPeriodLabel);
        _topFlowPanel.Controls.Add(_rpiSourceLabel);

        _runDateLabel.AutoSize = true;
        _runDateLabel.Margin = new Padding(0, 6, 24, 0);
        _runDateLabel.Text = "Run date:";

        _rpiLabel.AutoSize = true;
        _rpiLabel.Margin = new Padding(0, 6, 4, 0);
        _rpiLabel.Text = "RPI %:";

        _rpiNumericUpDown.DecimalPlaces = 1;
        _rpiNumericUpDown.Increment = 0.1m;
        _rpiNumericUpDown.Minimum = -50m;
        _rpiNumericUpDown.Maximum = 100m;
        _rpiNumericUpDown.Width = 70;
        _rpiNumericUpDown.Margin = new Padding(0, 3, 16, 0);
        _rpiNumericUpDown.TextAlign = HorizontalAlignment.Right;

        _rpiPeriodLabel.AutoSize = true;
        _rpiPeriodLabel.Margin = new Padding(0, 6, 24, 0);

        _rpiSourceLabel.AutoSize = true;
        _rpiSourceLabel.Margin = new Padding(0, 6, 0, 0);
        _rpiSourceLabel.Font = new Font(_rpiSourceLabel.Font, FontStyle.Bold);

        _warningBanner.Dock = DockStyle.Bottom;
        _warningBanner.Height = 28;
        _warningBanner.TextAlign = ContentAlignment.MiddleLeft;
        _warningBanner.Padding = new Padding(8, 0, 0, 0);
        _warningBanner.BackColor = Color.LightGoldenrodYellow;
        _warningBanner.Visible = false;

        // Master/detail split
        _mainSplitContainer.Dock = DockStyle.Fill;
        _mainSplitContainer.Orientation = Orientation.Vertical;
        _mainSplitContainer.SplitterWidth = 6;

        // Left: candidates grid
        _approveColumn.HeaderText = "Approve";
        _approveColumn.Width = 60;
        _clientColumn.HeaderText = "Client";
        _clientColumn.Width = 160;
        _clientColumn.ReadOnly = true;
        _dueDateColumn.HeaderText = "Due date";
        _dueDateColumn.Width = 90;
        _dueDateColumn.ReadOnly = true;
        _statusColumn.HeaderText = "Status";
        _statusColumn.Width = 70;
        _statusColumn.ReadOnly = true;
        _netColumn.HeaderText = "Net";
        _netColumn.Width = 80;
        _netColumn.ReadOnly = true;
        _netColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        _totalColumn.HeaderText = "Total";
        _totalColumn.Width = 80;
        _totalColumn.ReadOnly = true;
        _totalColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

        _candidatesGrid.Dock = DockStyle.Fill;
        _candidatesGrid.BackgroundColor = SystemColors.Window;
        _candidatesGrid.AllowUserToAddRows = false;
        _candidatesGrid.AllowUserToDeleteRows = false;
        _candidatesGrid.AllowUserToResizeRows = false;
        _candidatesGrid.RowHeadersVisible = false;
        _candidatesGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        _candidatesGrid.MultiSelect = false;
        _candidatesGrid.AutoGenerateColumns = false;
        _candidatesGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        _candidatesGrid.Columns.AddRange(_approveColumn, _clientColumn, _dueDateColumn, _statusColumn, _netColumn, _totalColumn);

        _mainSplitContainer.Panel1.Controls.Add(_candidatesGrid);

        // Right: detail panel
        _detailPanel.Dock = DockStyle.Fill;
        _detailPanel.Padding = new Padding(8);

        _addressGroupBox.Dock = DockStyle.Top;
        _addressGroupBox.Height = 150;
        _addressGroupBox.Text = "Client";
        _addressGroupBox.Controls.Add(_addressLayout);

        _addressLayout.Dock = DockStyle.Fill;
        _addressLayout.ColumnCount = 1;
        _addressLayout.RowCount = 6;
        _addressLayout.Padding = new Padding(8);
        for (var i = 0; i < 6; i++)
        {
            _addressLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        }

        _clientNameLabel.AutoSize = true;
        _clientNameLabel.Font = new Font(_clientNameLabel.Font, FontStyle.Bold);
        _contactNameLabel.AutoSize = true;
        _address1Label.AutoSize = true;
        _address2Label.AutoSize = true;
        _townPostcodeLabel.AutoSize = true;
        _emailLabel.AutoSize = true;

        _addressLayout.Controls.Add(_clientNameLabel, 0, 0);
        _addressLayout.Controls.Add(_contactNameLabel, 0, 1);
        _addressLayout.Controls.Add(_address1Label, 0, 2);
        _addressLayout.Controls.Add(_address2Label, 0, 3);
        _addressLayout.Controls.Add(_townPostcodeLabel, 0, 4);
        _addressLayout.Controls.Add(_emailLabel, 0, 5);

        _descriptionColumn.HeaderText = "Description";
        _descriptionColumn.ReadOnly = true;
        _descriptionColumn.FillWeight = 34;
        _qtyColumn.HeaderText = "Qty";
        _qtyColumn.ReadOnly = true;
        _qtyColumn.FillWeight = 10;
        _currentPriceColumn.HeaderText = "Current price";
        _currentPriceColumn.ReadOnly = true;
        _currentPriceColumn.FillWeight = 19;
        _currentPriceColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        _proposedPriceColumn.HeaderText = "Proposed price";
        _proposedPriceColumn.ReadOnly = false;
        _proposedPriceColumn.FillWeight = 19;
        _proposedPriceColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        _upliftColumn.HeaderText = "Uplift %";
        _upliftColumn.ReadOnly = true;
        _upliftColumn.FillWeight = 18;
        _upliftColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

        _itemsGrid.Dock = DockStyle.Fill;
        _itemsGrid.BackgroundColor = SystemColors.Window;
        _itemsGrid.AllowUserToAddRows = false;
        _itemsGrid.AllowUserToDeleteRows = false;
        _itemsGrid.RowHeadersVisible = false;
        _itemsGrid.AutoGenerateColumns = false;
        _itemsGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        _itemsGrid.Columns.AddRange(_descriptionColumn, _qtyColumn, _currentPriceColumn, _proposedPriceColumn, _upliftColumn);

        _totalsLayout.Dock = DockStyle.Bottom;
        _totalsLayout.Height = 84;
        _totalsLayout.ColumnCount = 2;
        _totalsLayout.RowCount = 3;
        _totalsLayout.Padding = new Padding(8);
        _totalsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
        _totalsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));

        var subtotalCaption = new Label { AutoSize = true, Text = "Subtotal:" };
        var vatCaption = new Label { AutoSize = true, Text = "VAT:" };
        var totalCaption = new Label { AutoSize = true, Text = "Total:", Font = new Font(_totalLabel.Font, FontStyle.Bold) };
        _subtotalLabel.AutoSize = true;
        _vatLabel.AutoSize = true;
        _totalLabel.AutoSize = true;
        _totalLabel.Font = new Font(_totalLabel.Font, FontStyle.Bold);

        _totalsLayout.Controls.Add(subtotalCaption, 0, 0);
        _totalsLayout.Controls.Add(_subtotalLabel, 1, 0);
        _totalsLayout.Controls.Add(vatCaption, 0, 1);
        _totalsLayout.Controls.Add(_vatLabel, 1, 1);
        _totalsLayout.Controls.Add(totalCaption, 0, 2);
        _totalsLayout.Controls.Add(_totalLabel, 1, 2);

        _detailPanel.Controls.Add(_itemsGrid);
        _detailPanel.Controls.Add(_totalsLayout);
        _detailPanel.Controls.Add(_addressGroupBox);

        _mainSplitContainer.Panel2.Controls.Add(_detailPanel);
        _mainSplitContainer.SplitterDistance = 480;

        // Bottom bar
        _bottomPanel.Dock = DockStyle.Bottom;
        _bottomPanel.Height = 96;
        _bottomPanel.Padding = new Padding(8);

        _createButton.Text = "Create approved invoices";
        _createButton.AutoSize = true;
        _createButton.Location = new Point(8, 8);

        _closeButton.Text = "Close";
        _closeButton.AutoSize = true;
        _closeButton.Location = new Point(220, 8);

        _openFolderButton.Text = "Open invoice folder";
        _openFolderButton.AutoSize = true;
        _openFolderButton.Location = new Point(320, 8);
        _openFolderButton.Visible = false;

        _buildProgressBar.Location = new Point(8, 40);
        _buildProgressBar.Width = 400;
        _buildProgressBar.Height = 20;

        _statusLabel.AutoSize = true;
        _statusLabel.Location = new Point(420, 44);

        _summaryTextBox.Multiline = true;
        _summaryTextBox.ReadOnly = true;
        _summaryTextBox.ScrollBars = ScrollBars.Vertical;
        _summaryTextBox.Location = new Point(8, 40);
        _summaryTextBox.Size = new Size(760, 50);
        _summaryTextBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        _summaryTextBox.Visible = false;

        _bottomPanel.Controls.Add(_createButton);
        _bottomPanel.Controls.Add(_closeButton);
        _bottomPanel.Controls.Add(_openFolderButton);
        _bottomPanel.Controls.Add(_buildProgressBar);
        _bottomPanel.Controls.Add(_statusLabel);
        _bottomPanel.Controls.Add(_summaryTextBox);

        // Form
        AutoScaleMode = AutoScaleMode.Dpi;
        ClientSize = new Size(1100, 700);
        MinimumSize = new Size(820, 480);
        Text = "Renewal invoices due for review";
        StartPosition = FormStartPosition.CenterScreen;

        Controls.Add(_mainSplitContainer);
        Controls.Add(_bottomPanel);
        Controls.Add(_topPanel);
    }
}
