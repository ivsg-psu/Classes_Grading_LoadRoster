function props = fcn_LoadRoster_SECURE_setPassword(props)
%fcn_LoadRoster_SECURE_setPassword   sets password field in props structure
% Only works if an authorized computer is listed. Set up so that it can be
% used via functions within _LOADROSTER_ repo, without sharing password.
%
% To develop a method to set passwords, see:
% https://www.mathworks.com/matlabcentral/answers/1672544-using-gmail-after-may-30-2022
% https://www.google.com/search?q=where+is+setting+for+less+secure+apps+in+gmail
