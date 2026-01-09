function fcn_LoadRoster_sendEmail( recipient, subject, body, attachments, varargin)
%% fcn_LoadRoster_sendEmail
% fcn_LoadRoster_sendEmail sends an email using MS Outlook.
% Allows users to specify body, attachments, etc.
%
% NOTE: The format of the function is similar to the SENDMAIL command.
% See:
% https://www.mathworks.com/matlabcentral/answers/94446-can-i-send-e-mail-through-matlab-using-microsoft-outlook
%
% FORMAT:
%
%      fcn_LoadRoster_sendEmail( recipient, subject, body, attachments, (figNum))
%
% INPUTS:
%
%      recipient: a string containing the email to send to
%
%      subject: a string containing the subject line
%
%      body: a string containing the body of the email
%
%      attachments: a cell array of strings that are paths to attachments
%
%      (optional inputs)
%
%      figNum: a figure number to plot results. If set to -1, skips any
%      input checking or debugging, no figures will be generated, and sets
%      up code to maximize speed.
%
% OUTPUTS:
%
%      (none)
%
% DEPENDENCIES:
%
%      fcn_DebugTools_checkInputsToFunctions
%
% EXAMPLES:
%
%      See the script:
%      script_test_fcn_LoadRoster_sendEmail
%      test suite.
%
% This function was written on 2026_01_09 by S. Brennan
% Questions or comments? sbrennan@psu.edu

% REVISION HISTORY:
%
% 2026_01_09 by Sean Brennan, sbrennan@psu.edu
% - wrote the code, using fcn_send+OutlookMail in AutoExam repo as starter

% TO-DO:
%
% 2026_01_09 by Sean Brennan, sbrennan@psu.edu
% - (fill in items here)



%% Debugging and Input checks

% Check if flag_max_speed set. This occurs if the figNum variable input
% argument (varargin) is given a number of -1, which is not a valid figure
% number.
MAX_NARGIN = 5; % The largest Number of argument inputs to the function
flag_max_speed = 0; % The default. This runs code with all error checking
if (nargin==MAX_NARGIN && isequal(varargin{end},-1))
    flag_do_debug = 0; % % % % Flag to plot the results for debugging
    flag_check_inputs = 0; % Flag to perform input checking
    flag_max_speed = 1;
else
    % Check to see if we are externally setting debug mode to be "on"
    flag_do_debug = 0; % % % % Flag to plot the results for debugging
    flag_check_inputs = 1; % Flag to perform input checking
    MATLABFLAG_LOADROSTER_FLAG_CHECK_INPUTS = getenv("MATLABFLAG_LOADROSTER_FLAG_CHECK_INPUTS");
    MATLABFLAG_LOADROSTER_FLAG_DO_DEBUG = getenv("MATLABFLAG_LOADROSTER_FLAG_DO_DEBUG");
    if ~isempty(MATLABFLAG_LOADROSTER_FLAG_CHECK_INPUTS) && ~isempty(MATLABFLAG_LOADROSTER_FLAG_DO_DEBUG)
        flag_do_debug = str2double(MATLABFLAG_LOADROSTER_FLAG_DO_DEBUG);
        flag_check_inputs  = str2double(MATLABFLAG_LOADROSTER_FLAG_CHECK_INPUTS);
    end
end

% flag_do_debug = 1;

if flag_do_debug % If debugging is on, print on entry/exit to the function
    st = dbstack; %#ok<*UNRCH>
    fprintf(1,'STARTING function: %s, in file: %s\n',st(1).name,st(1).file);
    debug_figNum = 999978; %#ok<NASGU>
else
    debug_figNum = []; %#ok<NASGU>
end

%% check input arguments?
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%   _____                   _
%  |_   _|                 | |
%    | |  _ __  _ __  _   _| |_ ___
%    | | | '_ \| '_ \| | | | __/ __|
%   _| |_| | | | |_) | |_| | |_\__ \
%  |_____|_| |_| .__/ \__,_|\__|___/
%              | |
%              |_|
% See: http://patorjk.com/software/taag/#p=display&f=Big&t=Inputs
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
if 0==flag_max_speed
    if flag_check_inputs
        % Are there the right number of inputs?
        narginchk(4,MAX_NARGIN);
    end
end

% Does user want to show the plots?
flag_do_plots = 0; % Default is to NOT show plots
if (0==flag_max_speed) && (MAX_NARGIN == nargin)
    temp = varargin{end};
    if ~isempty(temp) % Did the user NOT give an empty figure number?
        figNum = temp; %#ok<NASGU>
        flag_do_plots = 1;
    end
end

%% Main code
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%   __  __       _
%  |  \/  |     (_)
%  | \  / | __ _ _ _ __
%  | |\/| |/ _` | | '_ \
%  | |  | | (_| | | | | |
%  |_|  |_|\__,_|_|_| |_|
%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

server='smtp.gmail.com';
mail = 'prof.brennan@gmail.com';


% Create object and set parameters.
setpref('Internet','E_mail', mail);
setpref('Internet','SMTP_Server', server);
setpref('Internet','SMTP_Username', mail);
props = java.lang.System.getProperties;
props.setProperty( 'mail.smtp.auth', 'true' );
props.setProperty( 'mail.smtp.user', mail );
props.setProperty( 'mail.smtp.host', server );
props.setProperty( 'mail.smtp.port', '587' );
props.setProperty( 'mail.smtp.starttls.enable', 'true' );
props = fcn_LoadRoster_SECURE_setPassword(props); %#ok<NASGU>

% Send the email
sendmail(recipient, subject, body, attachments);


%% Any debugging?
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%   _____       _
%  |  __ \     | |
%  | |  | | ___| |__  _   _  __ _
%  | |  | |/ _ \ '_ \| | | |/ _` |
%  | |__| |  __/ |_) | |_| | (_| |
%  |_____/ \___|_.__/ \__,_|\__, |
%                            __/ |
%                           |___/
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
if flag_do_plots
    % Nothing to do here
end

if flag_do_debug
    fprintf(1,'ENDING function: %s, in file: %s\n\n',st(1).name,st(1).file);
end
end % Ends main function

%% Functions follow
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%   ______                _   _
%  |  ____|              | | (_)
%  | |__ _   _ _ __   ___| |_ _  ___  _ __  ___
%  |  __| | | | '_ \ / __| __| |/ _ \| '_ \/ __|
%  | |  | |_| | | | | (__| |_| | (_) | | | \__ \
%  |_|   \__,_|_| |_|\___|\__|_|\___/|_| |_|___/
%
% See: https://patorjk.com/software/taag/#p=display&f=Big&t=Functions
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%§
