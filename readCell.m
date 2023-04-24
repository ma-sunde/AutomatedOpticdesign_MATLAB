function [num_pos, param_pos] = readCell(cell_array, num, param)
    num_pos = -1;
    param_pos = -1;
    for n = 1:size(cell_array, 2)
        if cell_array{1,n} == num
            num_pos = n;
            break;
        end
    end

    for n = 1:size(cell_array, 1)
        if strcmp(cell_array{n,1}, param) == true
            param_pos = n;
            break;
        end
    end
    if num_pos == -1
        error("Surface number not found in cell array")
    elseif param_pos == -1
        error("Parameter not found in cell array")
    end
end