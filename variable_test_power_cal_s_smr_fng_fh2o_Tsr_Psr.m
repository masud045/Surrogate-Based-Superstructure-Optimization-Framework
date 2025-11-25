clc
clear

% Connect to Aspen
Aspen = actxserver('Apwn.Document.37.0');
pause(15);
Aspen.invoke('InitFromArchive2', 'C:\Users\mam00103\Documents\SMR\H2.apw');
pause(15);
Aspen.visible = 1;
pause(15);

% Set the number of samples for each variable
numSamples_FNG = 4;
numSamples_FH2O = 10;
numSamples_Tsr = 10;
numSamples_Psr = 2;

% Define the variable ranges
FNG_range = [80, 100]; %kmol/h
FH2O_range = [255, 310]; %kmol/h
Tsr_range = [900, 1000]; %C
Psr_range = [29.6, 38.2]; %atm

% Generate hypercube Latin samples
latinSamples_FNG = lhsdesign(numSamples_FNG, 1);
latinSamples_FH2O = lhsdesign(numSamples_FH2O, 1);
latinSamples_Tsr = lhsdesign(numSamples_Tsr, 1);
latinSamples_Psr = lhsdesign(numSamples_Psr, 1);

% Initialize arrays to store all combinations
totalSamples = numSamples_FNG * numSamples_FH2O * numSamples_Tsr * numSamples_Psr;
all_FNG_samples = zeros(totalSamples, 1);
all_FH2O_samples = zeros(totalSamples, 1);
all_Tsr_samples = zeros(totalSamples, 1);
all_Psr_samples = zeros(totalSamples, 1);
all_f = zeros(totalSamples, 1);
all_xh2 = zeros(totalSamples, 1);
net_energy = zeros(totalSamples, 1);

% Loop through the Latin Hypercube Samples and Aspen simulation
index = 1;
for k = 1:numSamples_FNG
    for l = 1:numSamples_FH2O
        for m = 1:numSamples_Tsr
            for n = 1:numSamples_Psr
                % Set the values for FNG, FH2O, Tsr, and Psr before each Aspen run
                Aspen.Application.Tree.FindNode('\Data\Streams\IN-NG\Input\TOTFLOW\MIXED').value = FNG_range(1) + latinSamples_FNG(k) * (FNG_range(2) - FNG_range(1));
                Aspen.Application.Tree.FindNode('\Data\Streams\IN-H2O\Input\TOTFLOW\MIXED').value = FH2O_range(1) + latinSamples_FH2O(l) * (FH2O_range(2) - FH2O_range(1));
                Aspen.Application.Tree.FindNode('\Data\Blocks\SR-REQ\Input\TEMP').value = Tsr_range(1) + latinSamples_Tsr(m) * (Tsr_range(2) - Tsr_range(1));
                Aspen.Application.Tree.FindNode('\Data\Blocks\SR-REQ\Input\PRES').value = Psr_range(1) + latinSamples_Psr(n) * (Psr_range(2) - Psr_range(1));
        
                % Run Aspen simulation
                Aspen.invoke('Run');
        
                % Retrieve the result and store it in arrays
                all_FNG_samples(index) = Aspen.Application.Tree.FindNode('\Data\Streams\IN-NG\Input\TOTFLOW\MIXED').value;
                all_FH2O_samples(index) = Aspen.Application.Tree.FindNode('\Data\Streams\IN-H2O\Input\TOTFLOW\MIXED').value;
                all_Tsr_samples(index) = Aspen.Application.Tree.FindNode('\Data\Blocks\SR-REQ\Input\TEMP').value;
                all_Psr_samples(index) = Aspen.Application.Tree.FindNode('\Data\Blocks\SR-REQ\Input\PRES').value;
                all_xh2(index) = Aspen.Application.Tree.FindNode('\Data\Streams\H2\Output\MOLEFRAC\MIXED\H2').value;
                all_f(index) = Aspen.Application.Tree.FindNode('\Data\Streams\H2\Output\TOT_FLOW').value;
                net_energy(index) = Aspen.Application.Tree.FindNode('\Data\Blocks\NG-COMP\Output\WNET').value + ...
                                    Aspen.Application.Tree.FindNode('\Data\Blocks\H2O-HX\Output\QNET').value + ...
                                    Aspen.Application.Tree.FindNode('\Data\Blocks\NG-HX\Output\QNET').value + ...
                                    Aspen.Application.Tree.FindNode('\Data\Blocks\H2O-COMP\Output\WNET').value + ...
                                    Aspen.Application.Tree.FindNode('\Data\Blocks\STEAM-HX\Output\QNET').value + ...
                                    Aspen.Application.Tree.FindNode('\Data\Blocks\PR-GIBBS\Output\QNET').value + ...
                                    Aspen.Application.Tree.FindNode('\Data\Blocks\HX-VAP\Output\QNET').value + ...
                                    Aspen.Application.Tree.FindNode('\Data\Blocks\SR-REQ\Output\QNET').value + ...
                                    Aspen.Application.Tree.FindNode('\Data\Blocks\HT-WGS\Output\QNET').value + ...
                                    Aspen.Application.Tree.FindNode('\Data\Blocks\HX-WGS\Output\QNET').value + ...
                                    Aspen.Application.Tree.FindNode('\Data\Blocks\LT-WGS\Output\QNET').value + ...
                                    Aspen.Application.Tree.FindNode('\Data\Blocks\HX-WGS-2\Output\QNET').value + ...
                                    Aspen.Application.Tree.FindNode('\Data\Blocks\H20FLASH\Output\QNET').value + ...
                                    Aspen.Application.Tree.FindNode('\Data\Blocks\CO2-SEP\Output\QCALC').value + ...
                                    Aspen.Application.Tree.FindNode('\Data\Blocks\HX-CO2\Output\QNET').value + ...
                                    Aspen.Application.Tree.FindNode('\Data\Blocks\METH\Output\QNET').value + ...
                                    Aspen.Application.Tree.FindNode('\Data\Blocks\HX-FL-2\Output\QNET').value + ...
                                    Aspen.Application.Tree.FindNode('\Data\Blocks\H2OFLASH\Output\QNET').value + ...
                                    Aspen.Application.Tree.FindNode('\Data\Blocks\DRYER\Output\QCALC').value + ...
                                    Aspen.Application.Tree.FindNode('\Data\Blocks\B1\Output\QNET').value + ...
                                    Aspen.Application.Tree.FindNode('\Data\Blocks\GAS-CMP1\Output\WNET').value; % Cal/s unit

                % Increment index
                index = index + 1;
            end
        end
    end
end

% Save the datasets to a MAT file
save('datasets1_power_cal_s_smr_fng_fh2o_Tsr_Psr.mat', 'all_FNG_samples', 'all_FH2O_samples','all_Tsr_samples', 'all_Psr_samples', 'net_energy');

% Display the xh2 matrix
disp('Resulting xh2 matrix:');
disp(all_xh2);

% Close Aspen
Aspen.invoke('Close');
