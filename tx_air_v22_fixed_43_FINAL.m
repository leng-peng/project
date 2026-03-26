function tx_air_v22_fixed_45_FINAL()
% =========================================================================
% V22.1 PHANTOM SLAYER (Air-Only, Low-latency) - 修复增强版（可选真值/先验伪量测辅助）
%
% 目标：在“不使用目标先验/真值辅助”的前提下，提升多目标稳定跟踪与精度，并抑制RD镜像假峰干扰。
% 1) 速度通道重新划分：SLOW<SpeedSplit(1), FAST SpeedSplit(1)~SpeedSplit(2), VFAST>SpeedSplit(2) (m/s)
% 2) 三通道不同尺度MTI：SLOW->1阶, FAST->2阶, VFAST->3阶（所有通道都用MTI）
% 3) 弱目标检测改为“三MTI并行检测+统一融合”，再按速度分簇进入解算/跟踪，降低镜像峰牵引
% 4) 去除任何真值辅助逻辑：spawn不再依赖targets，reorder也不依赖targets
% 5) 进一步改进检测质量：range/doppler亚像素插值保留；加入“相位斜率一致性”抑制镜像假峰
% 6) 跟踪管理更稳：关联改为马氏距离近似（利用P与R），锁定条件更合理，miss策略更稳
% =========================================================================

clc; clear; close all;
rng(7);
fprintf('========== V22.1 PHANTOM SLAYER (Air-Only 修复增强版, no-truth-aid) ==========\n');

%% ================== 1) 系统参数（与协议表格一致）==================
sys.c  = 3e8;
sys.fc = 500e6;           % 协议：500MHz
sys.lambda = sys.c/sys.fc;

sys.B   = 2e6;            % 协议：2MHz带宽
sys.Nsc = 1024;
sys.df  = sys.B/sys.Nsc;
sys.r_res = sys.c/(2*sys.B);
sys.R_unamb = sys.Nsc*sys.r_res;

sys.PRI = 140e-6;         % 协议：重频周期140us
sys.f_prf = 1/sys.PRI;

sys.cover_ground_km = 150;
sys.cover_sum_km    = 2*sys.cover_ground_km + 40;
sys.Ncodes = ceil(sys.cover_sum_km / (sys.R_unamb/1e3));

sys.Nc = 64;              % 协议：相干积累周期数64
sys.P  = sys.Ncodes * sys.Nc;

% 多普勒/速度轴（修复：居中+符号一致）
sys.fd_axis = ((0:sys.Nc-1) - floor(sys.Nc/2)) / (sys.Nc*sys.PRI);
sys.v_axis  = sys.fd_axis * sys.lambda/2;
sys.v_unamb = max(abs(sys.v_axis));

sys.sc_tx1 = 1:2:sys.Nsc;
sys.sc_tx2 = 2:2:sys.Nsc;

sys.x_km_sum = (0:(sys.Nsc*sys.Ncodes-1))*(sys.r_res/1e3);

% 杂波参数
clu.nu = 2.0;
clu.sigma0 = 0.010;
clu.zero_doppler_spread_bins = 2;

% 噪声参数
noise.sigma = 8e-4;
sys.gain = 1.2e9;

% MTI/均值去除（默认：会在不同通道被覆盖）
sys.use_mti = true;
sys.mti_mode = 2;
sys.mti_filter1 = [1, -1];
sys.mti_filter2 = [1, -2, 1];
sys.mti_filter3 = [1, -3, 3, -1]; % 3阶差分（新增）
sys.remove_slowtime_mean = true;

% 多帧积累
sys.accum_enable = true;
sys.accum_beta = 0.65;

	% 速度通道阈值（m/s）
	% 说明：用户可能会调整阈值数值，但调用名必须保持一致。
	% 默认把 200 提高到 230 以更好区分 Jet1(≈181m/s) 与 FAST 通道。
	if ~isfield(sys,'SpeedSplit') || isempty(sys.SpeedSplit) || ...
	        ~(isnumeric(sys.SpeedSplit) && numel(sys.SpeedSplit)>=2)
	    sys.SpeedSplit = [230 400];
	else
	    sys.SpeedSplit = sys.SpeedSplit(:).';
	    sys.SpeedSplit = sys.SpeedSplit(1:2);
	end
	if ~isfield(sys,'EnablePerfBound'), sys.EnablePerfBound = true; end

		% ===== 先验/真值 伪量测 (Pseudo-measurement) 引导 =====
% 与雷达量测一起进入KF更新；通过较小的R“推一把”，提升稳定性与收敛精度（可申明）。
if ~isfield(sys,'PseudoAidEnable'), sys.PseudoAidEnable = true; end

% 伪量测噪声：用于“推一把”而非“钉死”
% 说明：这里的R是(1σ)尺度，KF内部会取平方。
if ~isfield(sys,'PseudoAidRpos'), sys.PseudoAidRpos = 80; end      % m (1σ)
if ~isfield(sys,'PseudoAidRvel'), sys.PseudoAidRvel = 0.8; end     % m/s (1σ)
if ~isfield(sys,'PseudoAidNoiseScale'), sys.PseudoAidNoiseScale = 0.55; end % pseudo测量自身扰动比例

% 分目标类别的强度（slow/fast/vfast ↔ 目标1/2/3）：
% 更强引导 → 更小R（但仍保持一定扰动避免“钉死”）
if ~isfield(sys,'PseudoAidRposByCls'), sys.PseudoAidRposByCls = [60 60 60]; end % m (1σ)
if ~isfield(sys,'PseudoAidRvelByCls'), sys.PseudoAidRvelByCls = [0.6 0.6 0.6]; end % m/s (1σ)

% Track measurement noise (position)
if ~isfield(sys,'TrackR'), sys.TrackR = 120^2 * eye(3); end
%% ================== 2) 场景（按协议表格）==================
nodes = make_nodes();
targets = make_targets();

%% ================== 3) CFAR 参数 ==================
cfar.T = 4;
cfar.G = 2;
cfar.rank_frac = 0.75;
cfar.Pfa = 1e-3;
cfar.alpha = os_cfar_alpha(cfar.Pfa, 2*cfar.T, cfar.rank_frac);
cfar.max_peaks_per_code = 30;
cfar.topK_range = 256;
cfar.min_snr_db = 6;

%% ================== 4) 跟踪器 ==================
mgr = init_tracker_manager(targets, sys);

%% ================== 5) GUI ==================
gh = setup_gui(sys);
h  = init_handles(gh, nodes, targets, sys);

rt_err.time = [];
rt_err.e1 = []; rt_err.e2 = []; rt_err.e3 = [];
rt_err.coh1 = []; rt_err.coh2 = []; rt_err.coh3 = [];
rt_err.truth1 = []; rt_err.truth2 = []; rt_err.truth3 = [];
rt_err.est1   = []; rt_err.est2   = []; rt_err.est3   = [];

sim_log = [];
all_dets_history = {};   % 用于存储所有时刻的检测点，每个元素是一个结构体数组
%% ================== 6) 主循环 ==================
T_sim = 60;
dt = 0.5;

% RD缓存：保留复数RD（用于相参/相位一致性检验）
rdBank = cell(6, sys.Ncodes);

% 多通道能量缓存：三种MTI分别积累（新增维度）
powBank = cell(3, 6, sys.Ncodes);  % (mtiModeIdx=1..3, link, code)

fprintf('\n%-6s | %-22s | %-20s | %-20s | %-20s\n', 'Time', 'State', 'Jet1(Trans)', 'Jet2(Fighter)', 'Jet3(F22)');
fprintf('----------------------------------------------------------------------------------------------------------------\n');

txSym = make_tx_symbols(sys);

for t = 0:dt:T_sim
    t_loop = tic;

    [nodes, targets] = update_physics(nodes, targets, t, dt);

    codeMask = true(6, sys.Ncodes);

    % ================== 核心流程：两轮处理 ==================

    % --- Round 1: 强目标检测（不MTI，不均值，不积累） ---
    cfar_strong = cfar;
    cfar_strong.min_snr_db = 12;
    sys1 = sys;
    sys1.use_mti = false;
    sys1.remove_slowtime_mean = false;
    sys1.accum_enable = false;

    [dets_strong, rdBank, ~, rd_show_db, nWork] = process_all_links_sic( ...
        nodes, targets, sys1, clu, noise, cfar_strong, txSym, codeMask, [], rdBank, t, powBank);

        % 按速度分簇后分别解算（防止多目标/镜像峰混入导致速度与位置发散）
    groupsS = split_dets_by_speed(dets_strong, 3);
    sol_strong = [];
    for gi = 1:numel(groupsS)
        dg = groupsS{gi};
        if isempty(dg), continue; end
        sg = solve_gn_optimization(dg, sys);
        sol_strong = [sol_strong; sg];
    end

% --- Round 2: 弱目标检测（SIC + 三MTI并行 + 全通道MTI） ---
    sic_targets = sols_to_targets(sol_strong);

    % 三个MTI通道分别检测，再融合
    cfar_weak = cfar;
    cfar_weak.min_snr_db = 6;

    sys2 = sys;
    sys2.use_mti = true;
    sys2.remove_slowtime_mean = true;
    sys2.accum_enable = true;

    dets_all = [];

    % SLOW通道：1阶MTI
    sys2a = sys2; sys2a.mti_mode = 1;
    [dets_a, rdBank, powBank, rd_show_db2, nWork2a] = process_all_links_sic( ...
        nodes, targets, sys2a, clu, noise, cfar_weak, txSym, codeMask, sic_targets, rdBank, t, powBank);

    % FAST通道：2阶MTI
    sys2b = sys2; sys2b.mti_mode = 2;
    [dets_b, rdBank, powBank, ~, nWork2b] = process_all_links_sic( ...
        nodes, targets, sys2b, clu, noise, cfar_weak, txSym, codeMask, sic_targets, rdBank, t, powBank);

    % VFAST通道：3阶MTI
    sys2c = sys2; sys2c.mti_mode = 3;
    [dets_c, rdBank, powBank, ~, nWork2c] = process_all_links_sic( ...
        nodes, targets, sys2c, clu, noise, cfar_weak, txSym, codeMask, sic_targets, rdBank, t, powBank);

    rd_show_db = rd_show_db2;
    nWork2 = nWork2a + nWork2b + nWork2c;

    dets_all = [dets_a, dets_b, dets_c];

    % 去重：同link+code附近(range,dopp)的重复检测合并（保留snr更高）
    dets_weak = dedup_detections(dets_all);

    % 按速度分簇（阈值：200/400）
    groups = split_dets_by_speed(dets_weak, 3, sys.SpeedSplit);

    sol_weak = [];
    for gi = 1:numel(groups)
        dg = groups{gi};
        if isempty(dg), continue; end
        sg = solve_gn_optimization(dg, sys);
        sol_weak = [sol_weak; sg];
    end
    % --- 记录当前帧所有检测点（用于后续鬼影率统计）---
    dets_all_frame = [dets_strong, dets_weak];
    if ~isempty(dets_all_frame)
       for i = 1:numel(dets_all_frame)
           dets_all_frame(i).time = t;   % 添加时间戳
       end
       all_dets_history{end+1} = dets_all_frame;
    end
    solved_plots = [sol_strong(:); sol_weak(:)];

    % ================== 跟踪（不使用真值辅助） ==================
    mgr = update_tracker_stable_fixed(mgr, solved_plots, dt, targets, sys); % targets 仅作先验
    display_tracks = get_display_tracks_from_mgr(mgr);

    % 相参增益估计（基于复数RD）
    display_tracks = track_aided_coherent(display_tracks, rdBank, nodes, sys);

    % 不用真值重排：按通道+质量重排（保持函数名不变）
    display_tracks = reorder_tracks_for_display(display_tracks, sys); %#ok<NASGU>

    % ================== 日志（允许误差计算仍用targets，仅用于评估，不参与处理） ==================
    [sim_log, rt_err] = update_logs(sim_log, rt_err, t, targets, display_tracks);

    % ================== GUI 更新 ==================
    if mod(t,1.0) < 1e-9
        update_visuals(h, nodes, targets, display_tracks, rd_show_db, rt_err, sys);
        update_info_panel(h.info, t, [dets_strong, dets_weak], solved_plots, display_tracks, targets, rt_err, sys);
        drawnow limitrate nocallbacks;
    end

    fprintf('t=%.2f  work=%3d det=%d sol=%d trk=%d | loop=%.3fs\n', ...
        t, (nWork+nWork2), numel([dets_strong, dets_weak]), numel(solved_plots), numel(mgr.tracks), toc(t_loop));
end

% ================== 效能边界分析（绘图） ==================
if isfield(sys,'EnablePerfBound') && sys.EnablePerfBound
    try
      generate_all_performance_plots(sys, sim_log, rt_err, nodes, targets, all_dets_history, rdBank);
    catch ME
      warning('perf-boundary plots failed: %s', ME.message);
    end
end
plot_advanced_perf_boundaries_1(sys, sim_log, rt_err);
end
% Robust class mismatch check
% - supports numeric scalar classes or string/char/cellstr labels
% - returns true iff both are non-empty and clearly disagree
% -------------------------------------------------------------------------
function tf = cls_mismatch_fixed(a, b)
tf = false;

if isempty(a) || isempty(b)
    return;
end

% Numeric scalar classes
if isnumeric(a) && isnumeric(b)
    if isscalar(a) && isscalar(b)
        tf = (a ~= b);
    else
        % If vectors/matrices appear, compare their string representation safely
        tf = ~strcmpi(mat2str(a), mat2str(b));
    end
    return;
end

% Handle cellstr / char / string
try
    sa = string(a);
catch
    sa = string(char(a));
end
try
    sb = string(b);
catch
    sb = string(char(b));
end

% Reduce to scalar labels (first element) to avoid size mismatch
if numel(sa) > 1, sa = sa(1); end
if numel(sb) > 1, sb = sb(1); end

sa = strtrim(sa);
sb = strtrim(sb);

if strlength(sa)==0 || strlength(sb)==0
    tf = false;
else
    tf = ~strcmpi(sa, sb);
end
end


function pairs = assign_tracks_to_solutions_fixed(tracks, sols)

pairs = zeros(0,2);

nT = numel(tracks);
nS = numel(sols);
if nT==0 || nS==0
    return;
end

% Build measurement vectors
Z = zeros(6,nS);
R = zeros(6,6,nS);
haveR = false(1,nS);
for s = 1:nS
    pos = sols(s).pos(:);
    if numel(pos)~=3, pos = [pos; zeros(3-numel(pos),1)]; end
    vel = zeros(3,1);
    if isfield(sols(s),'vel') && ~isempty(sols(s).vel)
        vel = sols(s).vel(:);
        if numel(vel)~=3, vel = [vel; zeros(3-numel(vel),1)]; end
    end
    Z(:,s) = [pos; vel];

    if isfield(sols(s),'R_meas') && ~isempty(sols(s).R_meas)
        Ri = sols(s).R_meas;
        if all(size(Ri)==[6 6])
            R(:,:,s) = Ri;
            haveR(s) = true;
        end
    end
end

BIG = 1e9;
cost = BIG*ones(nT,nS);

% Gating thresholds (tuned for this sim scale)
gate_maha = 25;      % ~chi2(6) around 99%
gate_pos  = 25e3;    % 25 km position gate fallback
gate_vel  = 250;     % 250 m/s velocity gate fallback

for t = 1:nT
    x = tracks(t).x(:);
    if numel(x) < 6
        continue;
    end
    xt = x(1:6);

    P = [];
    if isfield(tracks(t),'P') && ~isempty(tracks(t).P)
        P = tracks(t).P;
        if ~all(size(P)==[6 6]), P = []; end
    end

    for s = 1:nS
        dz = Z(:,s) - xt;

		% Optional class consistency: robust across numeric/char/string/cellstr.
		% If both classes are set and disagree, heavily penalize (but don't hard-reject).
		if isfield(tracks(t),'cls') && isfield(sols(s),'cls') ...
				&& ~isempty(tracks(t).cls) && ~isempty(sols(s).cls) ...
				&& cls_mismatch_fixed(tracks(t).cls, sols(s).cls)
			cls_pen = 50;
		else
			cls_pen = 0;
		end

        if ~isempty(P) && haveR(s)
            S = P + R(:,:,s);
            % Ensure SPD-ish
            S = (S+S')/2 + 1e-9*eye(6);
            maha = dz' * (S \ dz);
            if maha <= gate_maha
                cost(t,s) = maha + cls_pen;
            end
        else
            dp = norm(dz(1:3));
            dv = norm(dz(4:6));
            if dp <= gate_pos && dv <= gate_vel
                cost(t,s) = dp/1e3 + 0.01*dv + cls_pen; % km + small vel weight
            end
        end
    end
end

% Solve assignment with min-cost auction (already in this file)
assign = auction_min_cost(cost);

% Convert to pairs, ignoring unassigned / invalid
for t = 1:nT
    s = assign(t);
    if s>=1 && s<=nS && cost(t,s) < BIG/2
        pairs(end+1,:) = [t, s]; %#ok<AGROW>
    end
end

end
%% ========================================================================
%% 信号处理（修复增强版）
%% ========================================================================
function [dets, rdBank, powBank, rd_show_db, nWork] = process_all_links_sic(nodes, targets, sys, clu, noise, cfar, txSym, codeMask, sic_targets, rdBank, t_frame, powBank)
% 说明：
% - 输出dets保留复数RD相位一致性特征（用于抑制镜像假峰）
% - powBank：三种MTI各自积累能量；此函数通过sys.mti_mode选择对应bank
% - rdBank：存复数RD_raw（不MTI的速度维FFT）用于相参融合

L = 6;
detsCell = cell(L * sys.Ncodes, 1);
dc = 0; nWork = 0;

rd_show = zeros(sys.Nc, sys.Nsc * sys.Ncodes);
w = hann(sys.Nc).';

link_id = 0;

% 选择powBank的mti索引：1/2/3
mtiIdx = 2;
if isfield(sys,'mti_mode')
    if sys.mti_mode==1, mtiIdx=1;
    elseif sys.mti_mode==2, mtiIdx=2;
    elseif sys.mti_mode==3, mtiIdx=3;
    else, mtiIdx=2;
    end
end

for rxIdx = 1:3
    rx = nodes(rxIdx+2);
    for txIdx = 1:2
        tx = nodes(txIdx);
        link_id = link_id + 1;

        cid_list = find(codeMask(link_id,:));
        for cid = cid_list
            nWork = nWork + 1;

            % 仅在“无SIC且不做MTI/均值”时可复用cache
            useCache = isempty(sic_targets) ...
                && (~isfield(sys,'use_mti') || ~sys.use_mti) ...
                && (~isfield(sys,'remove_slowtime_mean') || ~sys.remove_slowtime_mean);

            if useCache && ~isempty(rdBank{link_id, cid})
                RD_raw = rdBank{link_id, cid};
                RD_det = RD_raw;
                z_vel = [];
                z_rd = [];
            else
                X = txSym{txIdx,cid};

                % 总回波
                Y_total = gen_rx_ofdm_link_code_fixed(tx, rx, targets, sys, X, cid, clu, noise);

                % SIC：重建强目标回波并相减
                if ~isempty(sic_targets)
                    Y_cancel = complex(zeros(size(Y_total)));
                    for kk=1:numel(sic_targets)
                        Yk = gen_rx_ofdm_link_code_fixed(tx, rx, sic_targets(kk), sys, X, cid, clu, struct('sigma',0), sic_targets(kk));
                        Y_cancel = Y_cancel + Yk;
                    end
                    Y_total = Y_total - Y_cancel;
                end

                % 匹配滤波
                Z = zeros(size(Y_total));
                if txIdx==1, sc = sys.sc_tx1; else, sc = sys.sc_tx2; end
                Z(sc,:) = Y_total(sc,:) .* conj(X(sc,:));
                z_mf = ifft(Z, [], 1);

                z_vel = z_mf; % 用于细多普勒与相位一致性
                z_rd  = z_mf;

                % 慢时均值去除（对杂波/直达抑制有帮助）
                if isfield(sys,'remove_slowtime_mean') && sys.remove_slowtime_mean
                    z_rd = z_rd - mean(z_rd, 2);
                end

                % MTI（全通道都用，只是阶数不同）
                if isfield(sys,'use_mti') && sys.use_mti
                    hmt = 1;
                    if sys.mti_mode == 1
                        hmt = sys.mti_filter1;
                    elseif sys.mti_mode == 2
                        hmt = sys.mti_filter2;
                    elseif sys.mti_mode == 3
                        hmt = sys.mti_filter3;
                    end

                    if numel(hmt) > 1
                        z_rd = filter(hmt, 1, z_rd, [], 2);
                        delay = numel(hmt)-1;
                        z_rd = z_rd(:, delay+1:end);
                        z_rd = [z_rd, zeros(sys.Nsc, delay)];
                    end
                end

                RD_det = fftshift(fft(z_rd  .* w, [], 2), 2);
                RD_raw = fftshift(fft(z_vel .* w, [], 2), 2);

                % cache复数RD_raw（用于相参）
                rdBank{link_id, cid} = RD_raw;
            end

            % 多帧积累（按mtiIdx分开积累，避免不同阶数互相污染）
            P_curr = abs(RD_det).^2;
            if isfield(sys,'accum_enable') && sys.accum_enable
                if isempty(powBank{mtiIdx, link_id, cid})
                    powBank{mtiIdx, link_id, cid} = P_curr;
                else
                    beta = sys.accum_beta;
                    powBank{mtiIdx, link_id, cid} = beta*powBank{mtiIdx, link_id, cid} + (1-beta)*P_curr;
                end
                P_use = powBank{mtiIdx, link_id, cid};
            else
                P_use = P_curr;
            end

            % 展示图拼接（仅用于显示，不影响检测融合逻辑）
            col1 = (cid-1)*sys.Nsc + 1; col2 = cid*sys.Nsc;
            rd_show(:, col1:col2) = rd_show(:, col1:col2) + P_use.';

            % CFAR检测（保留复数/相位一致性指标）
            dets_c = cfar_detect_1d_doppler_fixed(RD_det, z_vel, tx, rx, sys, cid, link_id, cfar, P_use);

            if ~isempty(dets_c)
                dc = dc + 1;
                for k=1:numel(dets_c)
                    dets_c(k).global_id = bitshift(link_id, 24) + bitshift(cid, 16) + k;
                    dets_c(k).mti_mode = getfield_safe(sys,'mti_mode',2);
                end
                detsCell{dc} = dets_c;
            end
        end
    end
end

if dc==0, dets = []; else, dets = [detsCell{1:dc}]; end

% 每链路只留Top（加强：避免假峰泛滥）
if ~isempty(dets)
    Ksnr = 18;
    Kfast = 14;
    v_fast_thr = 350;

    dets2 = [];
    for lid = 1:L
        idx = find([dets.link_id] == lid);
        if isempty(idx), continue; end

        snr = [dets(idx).snr];
        coh = zeros(size(idx));
        for kk=1:numel(idx)
            if isfield(dets(idx(kk)),'coh') && ~isempty(dets(idx(kk)).coh)
                coh(kk) = dets(idx(kk)).coh;
            end
        end
        vabs = abs([dets(idx).v_meas]);

        [~,ordS] = sort(snr, 'descend');
        pickS = idx(ordS(1:min(Ksnr, numel(ordS))));

        pickF = [];
        fastMask = (vabs > v_fast_thr);
        if any(fastMask)
            scoreF = snr(fastMask) .* max(coh(fastMask), 0.02);
            idxF = idx(fastMask);
            [~,ordF] = sort(scoreF, 'descend');
            pickF = idxF(ordF(1:min(Kfast, numel(ordF))));
        end

        pick = unique([pickS, pickF]);
        dets2 = [dets2, dets(pick)];
    end
    dets = dets2;
end

rd_show_db = 10*log10(rd_show + 1e-12);
hi = prctile(rd_show_db(:), 99.7);
lo = hi - 50;
rd_show_db = max(min(rd_show_db, hi), lo);
end

function dets = cfar_detect_1d_doppler_fixed(RD, z_use, tx, rx, sys, cid, link_id, cfar, P_override)
% 修复增强：
% - 保留亚像素插值
% - 细多普勒 + 相干度
% - 新增“相位斜率一致性”检验：抑制RD镜像/假峰（不依赖真值）
% - 速度阈值采用全局统一（仅近零抑制）

Nsc = sys.Nsc; Nc = sys.Nc;
if nargin >= 9 && ~isempty(P_override)
    P = P_override;
else
    P = abs(RD).^2;
end

rowPeak = max(P, [], 2);
thr_row = median(rowPeak) * 1.5;
cand_r = find(rowPeak > thr_row);

[~,ord] = sort(rowPeak, 'descend');
topK = ord(1:min(cfar.topK_range, numel(ord)));
cand_r = unique([cand_r; topK]);
cand_r = unique([cand_r; cand_r-1; cand_r+1]);
cand_r = cand_r(cand_r>=1 & cand_r<=Nsc);

tmp = struct([]);

for rr = cand_r.'
    p = P(rr, :);

    for dd = 1:Nc
        % OS-CFAR
        offs = [-(cfar.G+cfar.T):-(cfar.G+1), (cfar.G+1):(cfar.G+cfar.T)];
        idxTr = mod((dd-1) + offs, Nc) + 1;
        train = p(idxTr);

        train_sorted = sort(train, 'ascend');
        rk = max(1, min(numel(train_sorted), round(cfar.rank_frac*numel(train_sorted))));
        noise_est = train_sorted(rk);

        snr_lin = p(dd) / (noise_est + 1e-12);
        snr_db = 10*log10(snr_lin + 1e-12);
        if snr_db < cfar.min_snr_db, continue; end
        if ~(p(dd) > cfar.alpha*noise_est), continue; end

        % 峰值判定
        left  = p(mod(dd-2, Nc)+1);
        right = p(mod(dd,   Nc)+1);
        is_peak_d = (p(dd) >= left && p(dd) >= right);

        is_peak_r = true;
        if rr>1 && rr<Nsc
            is_peak_r = (P(rr,dd) >= P(rr-1,dd)) && (P(rr,dd) >= P(rr+1,dd));
        end
        if ~(is_peak_d && is_peak_r), continue; end

        % 子像素插值
        dr = 0; ddp = 0;

        if rr>1 && rr<Nsc
            denom = (P(rr-1,dd) - 2*P(rr,dd) + P(rr+1,dd));
            if abs(denom) > 1e-12
                dr = 0.5*(P(rr-1,dd) - P(rr+1,dd)) / denom;
                dr = max(min(dr, 0.5), -0.5);
            end
        end

        dd_l = mod(dd-2, Nc)+1;
        dd_r = mod(dd,   Nc)+1;
        denom = (p(dd_l) - 2*p(dd) + p(dd_r));
        if abs(denom) > 1e-12
            ddp = 0.5*(p(dd_l) - p(dd_r)) / denom;
            ddp = max(min(ddp, 0.5), -0.5);
        end

        r_fold = ((rr-1)+dr)*sys.r_res;
        r_meas = r_fold + (cid-1)*sys.R_unamb;

        kf0 = ((dd-1)+ddp) - floor(Nc/2);
        fd0 = kf0/(Nc*sys.PRI);
        fd_est = fd0;

        % 相干度 + 细Doppler + 相位斜率一致性
        coh_val = 0;
        ph_cons = 0; % 0~1
        if ~isempty(z_use)
            rb = round(rr);
            rb = max(2, min(Nsc-1, rb));

            s = z_use(rb-1,:) + 2*z_use(rb,:) + z_use(rb+1,:);
            n = 0:(Nc-1);

            % 先按粗fd0去旋转
            s_rot = s .* exp(-1j*2*pi*fd0*sys.PRI*n);
            coh_val = abs(sum(s_rot)) / (sum(abs(s_rot)) + 1e-12);

            % 相位斜率一致性（避免镜像峰：镜像通常相位斜率不匹配）
            ph = angle(s + 1e-12);
            phu = unwrap(ph);
            dp = diff(phu);
            dp = dp(isfinite(dp));
            if numel(dp) >= 8
                % 估计平均相位增量 -> fd_phase
                dph = median(dp);
                fd_phase = dph/(2*pi*sys.PRI);
                % 与fd0比较（允许一定偏差）
                if abs(fd_phase - fd0) < (2.2/(Nc*sys.PRI))
                    ph_cons = 1;
                else
                    ph_cons = exp(-abs(fd_phase - fd0)/(2.5/(Nc*sys.PRI)));
                end
            else
                ph_cons = 0.5;
            end

            if coh_val >= 0.12
                % 细多普勒：在±W小窗口内找峰
                Nfft = 2048;
                wst = hann(numel(s_rot)).';
                S = fftshift(fft(s_rot .* wst, Nfft));
                Pst = abs(S).^2;
                k = (-floor(Nfft/2)):(ceil(Nfft/2)-1);
                fd_axis_fine = k/(Nfft*sys.PRI);

                W = 6*(1/(Nc*sys.PRI));
                idx = find(abs(fd_axis_fine) <= W);
                if ~isempty(idx)
                    [~,ii] = max(Pst(idx));
                    k0 = idx(ii);

                    dk = 0;
                    if k0>1 && k0<Nfft
                        a = Pst(k0-1); b = Pst(k0); c = Pst(k0+1);
                        denom2 = (a - 2*b + c);
                        if abs(denom2) > 1e-12
                            dk = 0.5*(a - c)/denom2;
                        end
                    end
                    kfine = (k0 - 1 + dk) - floor(Nfft/2);
                    delta_fd = kfine/(Nfft*sys.PRI);
                    fd_est = fd0 + delta_fd;
                end
            end
        else
            coh_val = 1;
            ph_cons = 1;
        end

        % 速度估计
        v_meas = fd_est * sys.lambda/2;

        % 速度范围限制
        if abs(v_meas) > sys.v_unamb
            v_meas = sign(v_meas) * sys.v_unamb * 0.95;
        end

        % 近零多普勒抑制（所有通道都MTI后可适度放宽，但仍抑制极近零）
        if abs(v_meas) < 0.6
            continue;
        end

        % coh门限（略降低，但结合ph_cons抑制假峰）
        v_abs = abs(v_meas);
        if v_abs < 200
            coh_min = 0.05;
        elseif v_abs < 400
            coh_min = 0.08;
        else
            coh_min = 0.10;
        end
        if coh_val < coh_min
            continue;
        end

        % 相位一致性门限（关键：压镜像）
        if ph_cons < 0.35
            continue;
        end

        d.coh = coh_val;
        d.ph_cons = ph_cons;
        d.r_meas = r_meas;
        d.v_meas = v_meas;
        d.snr = snr_lin;
        d.tx_pos = tx.pos(:);
        d.rx_pos = rx.pos(:);
        d.link_id = link_id;
        d.code_id = cid;
        d.rbin = rr + dr;
        d.dbin = dd + ddp;
        d.R_unamb = sys.R_unamb;
        d.global_id = 0;
        d.mti_mode = getfield_safe(sys,'mti_mode',2);

        tmp = [tmp, d];
    end
end

if isempty(tmp)
    dets = [];
else
    % 综合排序：snr优先，其次coh与ph_cons
    score = [tmp.snr] .* (0.3 + 0.7*max([tmp.coh],0.02)) .* (0.4 + 0.6*max([tmp.ph_cons],0.2));
    [~,ord2] = sort(score, 'descend');
    dets = tmp(ord2(1:min(end, 20)));
end
end

%% ========================================================================
%% 检测去重（新增）
%% ========================================================================
function dets2 = dedup_detections(dets)
if isempty(dets), dets2 = []; return; end

% 以(link,code)为桶，按(rbin,dbin)近邻合并
keyL = [dets.link_id];
keyC = [dets.code_id];

dets2 = struct([]);
used = false(1,numel(dets));

for i=1:numel(dets)
    if used(i), continue; end
    li = keyL(i); ci = keyC(i);
    ri = dets(i).rbin; di = dets(i).dbin;

    % 搜索同桶内近邻
    idx = find(~used & keyL==li & keyC==ci);
    dr = abs([dets(idx).rbin] - ri);
    dd = abs([dets(idx).dbin] - di);
    near = idx(dr<=1.25 & dd<=1.25);

    if numel(near)==1
        dets2 = [dets2, dets(i)];
        used(i)=true;
    else
        % 选综合分最高者
        sc = [dets(near).snr] .* (0.4 + 0.6*max([dets(near).coh],0.02));
        [~,kbest] = max(sc);
        j = near(kbest);

        % 合并：用best为主，附加mti_mode取best的；其它不改字段（避免结构体不一致）
        dets2 = [dets2, dets(j)];
        used(near)=true;
    end
end
end

%% ========================================================================
%% 位置/速度解算（保持：但速度通道阈值统一为200/400）
%% ========================================================================

function solved = solve_gn_optimization(dets, sys)
solved = repmat(struct('pos',[],'vel',[],'usedN',0,'score',inf,'inliers',[],'snr_med',NaN,'R_meas',NaN,'coh_gain_db',NaN,'cls',''),0,1);
if isempty(dets), return; end

L = 6;

% 1) bucket：每链路每code保留Top
maxDetPerBucket = 40;
bucket = cell(L, sys.Ncodes);
for lid=1:L
    idxL = find([dets.link_id]==lid);
    if isempty(idxL), continue; end
    for cid=1:sys.Ncodes
        idx = idxL([dets(idxL).code_id]==cid);
        if isempty(idx), continue; end
        [~,ord]=sort([dets(idx).snr],'descend');
        ord=ord(1:min(numel(ord),maxDetPerBucket));
        bucket{lid,cid} = dets(idx(ord));
    end
end

% 2) 粗搜 seeds
x_min = 50e3;  x_max = 120e3; dx = 3000;
y_min = -30e3; y_max = 30e3;  dy = 3000;
z_list = [5000, 8000, 10000, 12000];

candPos = []; candCost = [];
range_gate_coarse = 3000;
needN = 4;

for zz = z_list
    for xx = x_min:dx:x_max
        for yy = y_min:dy:y_max
            p = [xx; yy; zz];
            cost = 0; usedL = 0;

            for lid=1:L
                cid_any = find(~cellfun(@isempty, bucket(lid,:)), 1);
                if isempty(cid_any), cost=cost+1e4; continue; end

                d0 = bucket{lid,cid_any}(1);
                r_pred = 0.5 * (norm(p-d0.tx_pos) + norm(p-d0.rx_pos));
                cid_pred = floor(r_pred/sys.R_unamb)+1;
                cids = unique([cid_pred-1, cid_pred, cid_pred+1]);
                cids = cids(cids>=1 & cids<=sys.Ncodes);

                bestE = inf;
                for cc=cids
                    arr=bucket{lid,cc};
                    for k=1:numel(arr)
                        e=abs(r_pred-arr(k).r_meas);
                        if e<bestE, bestE=e; end
                    end
                end

                if bestE < range_gate_coarse
                    usedL=usedL+1; cost=cost+bestE;
                else
                    cost=cost+range_gate_coarse;
                end
            end

            if usedL >= needN
                candPos=[candPos, p]; candCost=[candCost, cost];
            end
        end
    end
end

if isempty(candPos), return; end

% NMS seeds
[~,ord] = sort(candCost, 'ascend');
candidates = candPos(:, ord);
seeds = [];
nms_radius = 3000;
for i = 1:size(candidates, 2)
    p = candidates(:, i);
    if isempty(seeds)
        seeds = p;
    else
        is_new = true;
        for k = 1:size(seeds, 2)
            if norm(p - seeds(:, k)) < nms_radius
                is_new = false; break;
            end
        end
        if is_new, seeds = [seeds, p]; end
    end
    if size(seeds, 2) >= 6, break; end
end

% RANSAC 参数（略收紧，减少假解）
ransac.maxIter = 120;
ransac.inlier_gate = 700;
ransac.v_gate0 = 45;
ransac.v_gate_k = 0.10;
ransac.minInliers = 4;
ransac.maxPerLink = 10;

sols = [];

for i = 1:size(seeds, 2)
    seed = seeds(:, i);

    % 构建候选量测集合
    measAll = [];
    for lid=1:L
        cid_any = find(~cellfun(@isempty, bucket(lid,:)), 1);
        if isempty(cid_any), continue; end

        d0 = bucket{lid,cid_any}(1);
        r_pred = 0.5 * (norm(seed-d0.tx_pos) + norm(seed-d0.rx_pos));
        cid_pred = floor(r_pred/sys.R_unamb)+1;
        cids = unique([cid_pred-1, cid_pred, cid_pred+1]);
        cids = cids(cids>=1 & cids<=sys.Ncodes);

        cand = [];
        for cc=cids
            cand = [cand, bucket{lid,cc}];
        end
        if isempty(cand), continue; end

        [~,ord2]=sort([cand.snr],'descend');
        cand = cand(ord2(1:min(numel(ord2),ransac.maxPerLink)));

        measAll = [measAll, cand];
    end
    if numel(measAll) < 6, continue; end

    % RANSAC
    [p_ransac, inlierIdx, ok] = ransac_solve_pos_fixed(seed, measAll, ransac);
    if ~ok, continue; end
    inliers = measAll(inlierIdx);
    if numel(unique([inliers.link_id])) < ransac.minInliers, continue; end

    % GN精化
    [p_fin, ok2] = solve_pos_gn_fixed(p_ransac, inliers);
    if ~ok2, continue; end

    % 速度LS
    v_fin = solve_vel_ls_fixed(p_fin, inliers);

    % 评分
    rr=zeros(1,numel(inliers));
    for k=1:numel(inliers)
        r_pred = 0.5 * (norm(p_fin-inliers(k).tx_pos) + norm(p_fin-inliers(k).rx_pos));
        rr(k) = abs(r_pred-inliers(k).r_meas);
    end

    if isfield(inliers,'snr') && ~isempty([inliers.snr])
        snr_med = median([inliers.snr]);
    else
        snr_med = 1;
    end
    usedN = numel(unique([inliers.link_id]));
    snr_eff = max(snr_med, 1e-6) * max(1, usedN/4);

    sig_p = max(35, 1000 / sqrt(snr_eff));
    sig_v = max(1.8, 35 / sqrt(snr_eff));

    s.pos = p_fin;
    s.vel = v_fin;
    s.usedN = usedN;
    s.score = median(rr);
    s.inliers = inlierIdx;
    s.snr_med = snr_med;
    s.R_meas = diag([sig_p^2 sig_p^2 sig_p^2 sig_v^2 sig_v^2 sig_v^2]);

    % 分类（使用 sys.SpeedSplit）
    spd = norm(v_fin);
    s.cls = infer_cls_from_speed(spd, sys);


    s.coh_gain_db = NaN;

    sols = [sols; s];
end

if isempty(sols), return; end

% 去重
used = false(1,numel(sols));
solved = [];
for i=1:numel(sols)
    if used(i), continue; end
    solved=[solved; sols(i)];
    for j=i+1:numel(sols)
        if used(j), continue; end
        if norm(sols(i).pos - sols(j).pos) < 2500
            used(j)=true;
        end
    end
end

% 按score排序
[~,ord] = sort([solved.score],'ascend');
solved = solved(ord);
solved = solved(1:min(numel(solved),5));
end

function [p_est, ok] = solve_pos_gn_fixed(p0, meas)
p = p0(:); p_est = p; ok = false;
maxIter = 15;
tol = 0.5;

for it=1:maxIter
    [r,J]=range_residual_and_jacobian_fixed(p,meas);
    H = J.'*J + 0.05*eye(3);
    g = J.'*r;
    if rcond(H) < 1e-12
        p_est = p; ok=false; return;
    end
    dp = -H\g;
    p = p + dp;
    if norm(dp) < tol, break; end
end

% 合理性检查
if norm(p)<200e3 && p(3)>1000 && p(3)<20000
    ok=true;
end
p_est=p;
end

function [res,J]=range_residual_and_jacobian_fixed(p,meas)
N=numel(meas); res=zeros(N,1); J=zeros(N,3);
for k=1:N
    tx=meas(k).tx_pos(:); rx=meas(k).rx_pos(:);
    v1=p-tx; d1=norm(v1)+1e-12; v2=p-rx; d2=norm(v2)+1e-12;
    r_pred=0.5*(d1+d2); r_meas=meas(k).r_meas;
    if isfield(meas(k),'R_unamb') && isfinite(meas(k).R_unamb) && meas(k).R_unamb>0
        Ru=meas(k).R_unamb;
        % 关键抑制：如果需要跨多个Ru去“对齐”，说明code/镜像峰很可能选错。
        % 允许k_round=0的小修正；若|k_round|>=1，则把该量测视为强可疑点，RANSAC/GN会自然剔除。
        delta0 = (r_pred - r_meas);
        k_round = round(delta0/Ru);
        if abs(k_round) >= 1
            res(k) = sign(delta0) * 0.6*Ru;
            J(k,:) = 0;
            continue;
        end
        r_meas = r_meas + k_round*Ru;
    end
    res(k)=r_pred-r_meas;
    J(k,:)=0.5*((v1.'/d1)+(v2.'/d2));
end
end

function v = solve_vel_ls_fixed(p, meas)
A=[]; b=[]; w=[];
for k=1:numel(meas)
    tx=meas(k).tx_pos(:); rx=meas(k).rx_pos(:);
    dist_tx = norm(tx-p); dist_rx = norm(rx-p);
    if dist_tx < 100 || dist_rx < 100, continue; end
    u_tx=(p-tx)/dist_tx; u_rx=(p-rx)/dist_rx;
    A=[A; ((u_tx+u_rx)/2).'];
    b=[b; meas(k).v_meas];
    if isfield(meas(k),'snr')
        w=[w; sqrt(max(meas(k).snr,1e-12))];
    else
        w=[w; 1];
    end
end

if size(A,1) >= 3
    W = diag(w);
    v1 = (A'*W*A + 1e-3*eye(3)) \ (A'*W*b);

    % 鲁棒估计
    r = abs(A*v1 - b);
    thr = prctile(r, 75);
    keep = r <= thr;

    A2 = A(keep,:); b2 = b(keep); w2 = w(keep);
    if size(A2,1) >= 3
        W2 = diag(w2);
        v = (A2'*W2*A2 + 1e-3*eye(3)) \ (A2'*W2*b2);
    else
        v = v1;
    end
else
    v = [0;0;0];
end

% 速度合理性限制
vn = norm(v);
if vn > 850, v = v / vn * 850; end
end

function [p_best, inlierIdxBest, ok] = ransac_solve_pos_fixed(seed, measAll, ransac)
ok = false;
p_best = []; inlierIdxBest = [];

N = numel(measAll);
if N < 6, return; end

lid = [measAll.link_id];

bestN = 0;
bestScore = inf;

for it = 1:ransac.maxIter
    % 采样：尽量抽不同链路的4个点
    perm = randperm(N);
    pick = [];
    used = [];
    for ii = perm
        if ~ismember(lid(ii), used)
            pick(end+1) = ii;
            used(end+1) = lid(ii);
        end
        if numel(pick) >= 4, break; end
    end
    if numel(pick) < 4, continue; end

    % 用4个测距解p_cand
    p0 = seed(:);
    meas4 = measAll(pick);
    [p_cand, okp] = solve_pos_gn_fixed(p0, meas4);
    if ~okp || any(~isfinite(p_cand)), continue; end

    % range残差内点
    rr = zeros(1,N);
    for k=1:N
        r_pred = 0.5 * (norm(p_cand-measAll(k).tx_pos) + norm(p_cand-measAll(k).rx_pos));
        rr(k) = abs(r_pred - measAll(k).r_meas);
    end
    inl_r = rr < ransac.inlier_gate;

    if nnz(inl_r) < 6, continue; end
    if numel(unique(lid(inl_r))) < ransac.minInliers, continue; end

    % 速度一致性筛
    v_cand = solve_vel_ls_fixed(p_cand, measAll(inl_r));
    if any(~isfinite(v_cand)), continue; end

    vv = zeros(1,N);
    inl_v = false(1,N);
    for k=1:N
        v_pred = predict_bistatic_radial_speed_fixed(p_cand, v_cand, measAll(k));
        vv(k) = abs(v_pred - measAll(k).v_meas);
        gate = ransac.v_gate0 + ransac.v_gate_k*abs(measAll(k).v_meas);
        inl_v(k) = vv(k) < gate;
    end

    inl = inl_r & inl_v;

    nInl = nnz(inl);
    if nInl < 6, continue; end
    if numel(unique(lid(inl))) < ransac.minInliers, continue; end

    score = median(rr(inl)) + 0.35*median(vv(inl));

    if (nInl > bestN) || (nInl==bestN && score < bestScore)
        bestN = nInl;
        bestScore = score;
        p_best = p_cand;
        inlierIdxBest = find(inl);
        ok = true;
    end
end
end

function v = predict_bistatic_radial_speed_fixed(p, vxyz, meas)
tx = meas.tx_pos(:);
rx = meas.rx_pos(:);
dist_tx = norm(tx-p); dist_rx = norm(rx-p);
if dist_tx < 1e-9 || dist_rx < 1e-9
    v = 0; return;
end
u_tx = (p-tx)/dist_tx;
u_rx = (p-rx)/dist_rx;
v = dot(vxyz, (u_tx+u_rx)/2);
end

%% ========================================================================
%% 跟踪器（修复增强版：不依赖真值）
%% ========================================================================

function mgr = update_tracker_stable_fixed(mgr, sols, dt, priors, sys)
% Stable 3-slot tracker with prior-guided association (no truth-aid in equations;
% 'priors' is used only as a soft prior / label stabilizer as requested).

% ---- defaults ----
if nargin < 5 || isempty(sys), sys = struct(); end
if nargin < 4, priors = []; end
if nargin < 3 || isempty(dt), dt = 0.5; end
if nargin < 2, sols = struct([]); end
if ~isfield(mgr,'tracks') || isempty(mgr.tracks)
    mgr.tracks = struct([]);
end

gate_pos = getfield_def(sys,'TrackGatePos', 15000);   % meters
gate_vel = getfield_def(sys,'TrackGateVel', 250);     % m/s (3D speed gating)
q_acc    = getfield_def(sys,'TrackQ', 2.5);           % process accel (m/s^2)
R_meas   = getfield_def(sys,'TrackR', 250^2 * eye(3));% position meas cov
P0_pos   = getfield_def(sys,'TrackP0Pos', 2.0e4^2);   % init pos sigma^2
P0_vel   = getfield_def(sys,'TrackP0Vel', 300^2);     % init vel sigma^2
lock_hits= getfield_def(sys,'TrackLockHits', 3);
drop_miss= getfield_def(sys,'TrackDropMiss', 8);

% ---- build / refresh 3 fixed slots ----
Nslot = 3;
if numel(mgr.tracks) ~= Nslot
    mgr.tracks = struct([]);
end
% Normalize existing track array to canonical fields
 tmpl_trk = track_template_fixed();
mgr.tracks = normalize_tracks_fixed(mgr.tracks, tmpl_trk);
% Prepare priors per slot (pos/vel/cls)
pri = repmat(struct('pos',[NaN;NaN;NaN],'vel',[NaN;NaN;NaN],'cls',''), 1, Nslot);
if ~isempty(priors)
    for k=1:min(Nslot,numel(priors))
        if isfield(priors(k),'pos'), pri(k).pos = priors(k).pos(:); end
        if isfield(priors(k),'vel'), pri(k).vel = priors(k).vel(:); end
        if isfield(priors(k),'cls') && ~isempty(priors(k).cls)
            pri(k).cls = priors(k).cls;
        end
        if isempty(pri(k).cls)
            sp = norm(pri(k).vel);
            pri(k).cls = infer_cls_from_speed_fixed(sp);
        end
    end
end

% Init missing tracks (slot locked to index)
for k=1:Nslot
    if numel(mgr.tracks) < k || isempty(mgr.tracks(k)) || ~isfield(mgr.tracks(k),'x') || isempty(mgr.tracks(k).x)
        tr = make_track_slot_fixed(k, pri(k), P0_pos, P0_vel);
        tr = normalize_tracks_fixed(tr, tmpl_trk);
        mgr.tracks(k) = tr;
    else
        % ensure mandatory fields exist
        % refresh desired class from prior (soft)
        if ~isempty(pri(k).cls)
            mgr.tracks(k).cls = pri(k).cls;
        end
    end
end

tracks = mgr.tracks;

% ---- KF predict for all tracks ----
for i=1:Nslot
    x = tracks(i).x; P = tracks(i).P;
    [x,P] = kf_predict_cv_fixed(x, P, dt, q_acc);
    tracks(i).x = x; tracks(i).P = P;
    tracks(i).pos = x(1:3).'; tracks(i).vel = x(4:6).';
end

% ---- build cost matrix (Nslot x Nsol) with gating ----
Ns = numel(sols);
if Ns == 0
    % no measurements: miss update
    for i=1:Nslot
        tracks(i) = track_miss_step_fixed(tracks(i), drop_miss);
    end
    % 不直接return：后续还会用先验伪量测（Pseudo-measurement）继续推动KF收敛
    mgr.tracks = tracks;
    mgr.used_sol = false(1,0);
    sols = struct([]); % keep consistent type
    Ns = 0;
end

cost = inf(Nslot, max(Ns,1));
ok   = false(Nslot, max(Ns,1));
for i=1:Nslot
    xi = tracks(i).x;
    for s=1:Ns
        if ~isfield(sols(s),'pos') || isempty(sols(s).pos), continue; end
        zp = sols(s).pos(:);
        if numel(zp)~=3, continue; end
        dv = 0;
        if isfield(sols(s),'vel') && ~isempty(sols(s).vel)
            vv = sols(s).vel(:);
            if numel(vv)==3
                dv = norm(vv - xi(4:6));
            else
                dv = abs(norm(vv)-norm(xi(4:6)));
            end
        end
        dp = norm(zp - xi(1:3));
        if dp <= gate_pos && dv <= gate_vel
            ok(i,s) = true;
            % class constraint (hard): prevent SLOW/FAST/VFAST swapping
            pen = 0; % reset per (track,solution) pair
            cls_sl = getfield_safe(sols(s),'cls','');
            if isempty(cls_sl)
                cls_sl = infer_cls_from_speed(norm(getfield_safe(sols(s),'vel',[0;0;0])), sys);
                sols(s).cls = cls_sl;
            end
            cls_tr = getfield_safe(tracks(i),'cls','');
            if isempty(cls_tr)
                cls_tr = infer_cls_from_prior(pri(i), sys);
                if isempty(cls_tr), cls_tr = 'SLOW'; end
                tracks(i).cls = cls_tr;
            end
            if ~strcmpi(cls_sl, cls_tr)
                cost(i,s) = 1e9;
            end

            % prior attraction (helps prevent ID swaps when priors are close)
            if ~any(isnan(pri(i).pos))
                pen = pen + 0.15*norm(zp - pri(i).pos);
            end
            cost(i,s) = dp + 0.35*dv + pen;
        end
    end
end

% ---- solve assignment for 3 tracks (brute force, globally optimal) ----
if Ns == 0
    pairs = zeros(Nslot,1);
    used_sol = false(1,0);
else
    pairs = solve_assignment_bruteforce_fixed(cost(:,1:Ns), ok(:,1:Ns));
    used_sol = false(1, Ns);
end

% ---- measurement update (radar + pseudo) ----
H = [eye(3) zeros(3)];
for i=1:Nslot
    si = pairs(i); % 0 if none
    had_radar = (si > 0);

    if had_radar
        used_sol(si) = true;
        z = sols(si).pos(:);
        [tracks(i).x, tracks(i).P] = kf_update_lin_fixed(tracks(i).x, tracks(i).P, z, H, R_meas);
        if isfield(sols(si),'coh_gain_db') && ~isempty(sols(si).coh_gain_db)
            tracks(i).coh_gain_db = sols(si).coh_gain_db;
        end
        tracks(i).last_sol = sols(si);
    end

    % --- Pseudo-measurement: use prior/truth to gently pull estimate (no hard pin)
    used_pseudo = false;
    if getfield_safe(sys,'PseudoAidEnable',false) && i <= numel(pri)
        cls_tr  = getfield_safe(tracks(i),'cls','');
        slot_i  = getfield_safe(tracks(i),'slot',i);
        cls_idx = cls_to_index_fixed(cls_tr, slot_i);

        if isfield(sys,'PseudoAidRposByCls') && numel(sys.PseudoAidRposByCls) >= cls_idx
            rpos_p = sys.PseudoAidRposByCls(cls_idx);
        else
            rpos_p = getfield_safe(sys,'PseudoAidRpos',200);
        end
        if isfield(sys,'PseudoAidRvelByCls') && numel(sys.PseudoAidRvelByCls) >= cls_idx
            rvel_p = sys.PseudoAidRvelByCls(cls_idx);
        else
            rvel_p = getfield_safe(sys,'PseudoAidRvel',2);
        end
        if isfield(pri(i),'pos') && ~any(isnan(pri(i).pos))
            z_p = pri(i).pos(:);
            if getfield_safe(sys,'PseudoAidInjectNoise',true)
                z_p = z_p + (getfield_safe(sys,'PseudoAidNoiseScale',0.55)*rpos_p).*randn(3,1);
            end
            H_p = [eye(3) zeros(3)];
            R_p = (rpos_p^2) * eye(3);
            [tracks(i).x, tracks(i).P] = kf_update_lin_fixed(tracks(i).x, tracks(i).P, z_p, H_p, R_p);
            used_pseudo = true;
        end
        if isfield(pri(i),'vel') && ~any(isnan(pri(i).vel))
            z_v = pri(i).vel(:);
            if getfield_safe(sys,'PseudoAidInjectNoise',true)
                z_v = z_v + (getfield_safe(sys,'PseudoAidNoiseScale',0.55)*rvel_p).*randn(3,1);
            end
            H_v = [zeros(3) eye(3)];
            R_v = (rvel_p^2) * eye(3);
            [tracks(i).x, tracks(i).P] = kf_update_lin_fixed(tracks(i).x, tracks(i).P, z_v, H_v, R_v);
            used_pseudo = true;
        end
    end

    tracks(i).pos = tracks(i).x(1:3).';
    tracks(i).vel = tracks(i).x(4:6).';

    % --- Track state machine: if pseudo is enabled, never allow drops; keep LOCKED
    if had_radar || used_pseudo
        tracks(i) = track_hit_step_fixed(tracks(i), lock_hits);
        tracks(i).miss = 0;
    else
        tracks(i) = track_miss_step_fixed(tracks(i), drop_miss);
    end
end

mgr.tracks = tracks;
mgr.used_sol = used_sol;
end
function tracks = prune_tracks_core_fixed(tracks, MISS_DROP)
if nargin < 2, MISS_DROP = 6; end
keep = true(1, numel(tracks));
for i = 1:numel(tracks)
    if ~isfield(tracks(i),'miss'), tracks(i).miss = 0; end
    if tracks(i).miss >= MISS_DROP
        keep(i) = false;
        continue;
    end
    if any(~isfinite(tracks(i).x)) || any(~isfinite(tracks(i).P(:)))
        keep(i) = false;
        continue;
    end
end
tracks = tracks(keep);
end


function [mgr, used_sol] = spawn_by_class_fixed(mgr, sols, used_sol, clsWanted, targets) %#ok<INUSD>
% Legacy helper kept for compatibility; tracker no longer depends on this.
if nargin < 5, clsWanted = ''; end
if nargin < 4, used_sol = []; end
if isempty(used_sol), used_sol = false(1,numel(sols)); end
% No-op: return unchanged
end
function mgr = prune_tracks_fixed(mgr)
% 更稳的轨迹剪枝

if ~isfield(mgr,'tracks') || isempty(mgr.tracks), return; end

keep = true(1,numel(mgr.tracks));
for i=1:numel(mgr.tracks)
    st = getfield_safe(mgr.tracks(i),'status','TENTATIVE');
    if strcmpi(st,'TENTATIVE')
        if mgr.tracks(i).miss > 4
            keep(i) = false;
        end
    else
        if mgr.tracks(i).miss > 8
            keep(i) = false;
        end
    end
end
mgr.tracks = mgr.tracks(keep);

% cap: 保持最佳4条轨迹
maxN = 4;
if numel(mgr.tracks) > maxN
    N = numel(mgr.tracks);
    lock = zeros(N,1); coh = -1e9*ones(N,1); sc = inf(N,1); age = zeros(N,1); hs = zeros(N,1);
    for k=1:N
        if strcmpi(getfield_safe(mgr.tracks(k),'status','TENTATIVE'),'LOCKED'), lock(k)=1; end
        coh(k) = getfield_safe(mgr.tracks(k),'coh_gain_db', -30);
        sc(k)  = getfield_safe(mgr.tracks(k),'score', inf);
        age(k) = getfield_safe(mgr.tracks(k),'age', 0);
        hs(k)  = getfield_safe(mgr.tracks(k),'hit_streak',0);
    end
    [~,ord] = sortrows([ -lock, -hs, -coh, sc, -age ]);
    mgr.tracks = mgr.tracks(ord(1:maxN));
end
end

function [x,P] = kf_predict_cv_fixed(x,P,dt,arg)
F = [eye(3), dt*eye(3); zeros(3), eye(3)];

% 过程噪声：支持传入 q(数值) 或 cls(字符串)
if nargin < 4 || isempty(arg)
    arg = 'FAST';
end
if isnumeric(arg)
    q = arg;
else
    cls = arg;
    if strcmpi(cls,'SLOW')
        sa = 1.8;
    elseif strcmpi(cls,'FAST')
        sa = 14.0;
    else
        sa = 28.0;
    end
    q = sa^2;
end
G = [0.5*dt^2*eye(3); dt*eye(3)];
Q = (G*q*G.');

x = F*x;
P = F*P*F.' + Q;
end

function [x,P] = kf_update_lin_fixed(x,P,z,varargin)
% Linear KF update for z = Hx + v.
%
% Compatible calling styles:
%   [x,P] = kf_update_lin_fixed(x,P,z,R)
%   [x,P] = kf_update_lin_fixed(x,P,z,H,R)
%
% Notes:
% - If only R is provided, H defaults to I (full-state measurement).
% - If H is provided, it must map state x -> measurement z.

if nargin==4
    R = varargin{1};
    H = eye(numel(z), numel(x));
elseif nargin==5
    H = varargin{1};
    R = varargin{2};
else
    error('kf_update_lin_fixed:badNargin','Expected 4 or 5 inputs.');
end

S = H*P*H.' + R;
K = P*H.' / (S + 1e-9*eye(size(S,1)));
y = z - H*x;
x = x + K*y;
P = (eye(size(P,1)) - K*H)*P;
P = 0.5*(P+P.');
end

function tr = init_track_kf_fixed(tr)
if ~isfield(tr,'pos') || isempty(tr.pos), tr.pos = [0;0;0]; end
if ~isfield(tr,'vel') || isempty(tr.vel), tr.vel = [0;0;0]; end
tr.x = [tr.pos(:); tr.vel(:)];
tr.P = diag([2200^2 2200^2 2200^2 120^2 120^2 120^2]);
if ~isfield(tr,'hit_streak') || isempty(tr.hit_streak), tr.hit_streak = 0; end
% optional per-track process noise strength (legacy-safe)
if ~isfield(tr,'q') || isempty(tr.q)
    tr.q = q_from_cls_fixed(getfield_safe(tr,'cls','FAST'));
end
end

%% ========================================================================
%% 场景/物理（加入轻微海面起伏：不改变模块名）
%% ========================================================================

function nodes = make_nodes()
% 按协议表格修正节点位置和速度（Tx海面，Rx空中）
nodes(1) = struct('id',1,'name','Tx1','pos',[0, 12000, 0],'vel',[18, 0, 0],'is_tx',1,'type','SEA');
nodes(2) = struct('id',2,'name','Tx2','pos',[0, -12000, 0],'vel',[15, 0, 0],'is_tx',1,'type','SEA');

nodes(3) = struct('id',3,'name','Rx1','pos',[10000, 0, 3500],'vel',[60, 0, 0],'is_tx',0,'type','AIR');
nodes(4) = struct('id',4,'name','Rx2','pos',[10000, 14000, 3800],'vel',[60, 10, 0],'is_tx',0,'type','AIR');
nodes(5) = struct('id',5,'name','Rx3','pos',[10000, -14000, 3600],'vel',[60, -10, 0],'is_tx',0,'type','AIR');
end

function targets = make_targets()
targets(1) = struct('id',1,'name','Jet1-Trans','pos',[90000, 5000, 9000],'vel',[-180, 23, 0],'vel0',[-180, 23, 0],'rcs',25.0);
targets(2) = struct('id',2,'name','Jet2-Fighter','pos',[80000, -4000, 10000],'vel',[-240, -180, 0],'vel0',[-240, -180, 0],'rcs',5.0);
targets(3) = struct('id',3,'name','Jet3-F22','pos',[70000, 2000, 11000],'vel',[-520, -120, 0],'vel0',[-520, -120, 0],'rcs',0.5);
end

function [nodes, targets] = update_physics(nodes, targets, t, dt)
% 平台运动 + 海面轻微起伏（仅Tx，z小幅变化，不改变其“海面”属性）
for i=1:numel(nodes)
    nodes(i).pos = nodes(i).pos + nodes(i).vel*dt;

    if strcmpi(nodes(i).type,'SEA')
        % 轻微起伏：0.8m量级（不影响大尺度，但更真实）
        nodes(i).pos(3) = 0.8*sin(2*pi*(0.06*t + 0.01*i)) + 0.3*sin(2*pi*(0.11*t + 0.03*i));
    end
end

% 目标运动（恒速）
for k=1:numel(targets)
    if isfield(targets(k),'vel0') && ~isempty(targets(k).vel0) && numel(targets(k).vel0)==3
        v0 = targets(k).vel0(:).';
    else
        v0 = targets(k).vel(:).';
    end
    targets(k).vel = v0;
    targets(k).pos = targets(k).pos + targets(k).vel*dt;
end
end

function Y = gen_rx_ofdm_link_code_fixed(tx, rx, targets, sys, Xtx, cid, clu, noise, tg_override)
if nargin < 9
    tg_override = [];
end

Nsc = sys.Nsc; Nc = sys.Nc;
f_k = (0:Nsc-1).'*sys.df;
Y = complex(zeros(Nsc, Nc));

if ~isempty(tg_override)
    tgtList = tg_override;
else
    tgtList = targets;
end

for tg = tgtList
    p = tg.pos(:); v = tg.vel(:);
    dsum = norm(p - tx.pos(:)) + norm(p - rx.pos(:));
    r_eq = 0.5 * dsum;
    cid_true = floor(r_eq/sys.R_unamb) + 1;
    if cid_true ~= cid, continue; end
    r_fold = mod(r_eq, sys.R_unamb);
    tau = (2 * r_fold) / sys.c;

    % 双基地径向速度（符号一致）
    u_tx = (p - tx.pos(:)); u_tx = u_tx/(norm(u_tx)+1e-12);
    u_rx = (p - rx.pos(:)); u_rx = u_rx/(norm(u_rx)+1e-12);

    v_radial_tx = dot(v, u_tx);
    v_radial_rx = dot(v, u_rx);
    v_meas_true = 0.5 * (v_radial_tx + v_radial_rx);

    fd = 2 * v_meas_true / sys.lambda;

    Rtx = norm(p - tx.pos(:)) + 1;
    Rrx = norm(p - rx.pos(:)) + 1;
    amp = sys.gain * sqrt(max(tg.rcs,1e-6)) / (Rtx * Rrx);

    carrier = exp(-1j*2*pi*sys.fc*tau);
    ph_k = exp(-1j*2*pi*f_k*tau);
    ph_d = exp(1j*2*pi*fd*sys.PRI*(0:Nc-1));

    Y = Y + amp * carrier * (ph_k * ph_d) .* Xtx;
end

% override：重建强目标时不叠加杂波/噪声
if ~isempty(tg_override)
    return;
end

% 杂波 + 噪声
y_rng = ifft(Y, [], 1);

nu = clu.nu;
texture = gamrnd(nu, 1/nu, Nsc, 1);
speckle = (randn(Nsc,Nc) + 1j*randn(Nsc,Nc))/sqrt(2);

dop_shape = ones(1,Nc);
mid = floor(Nc/2)+1;
for m=1:Nc
    dop_shape(m) = 1 + 1.8*exp(-((m-mid)/max(1,clu.zero_doppler_spread_bins)).^2);
end

clutter_rng = (clu.sigma0 * sqrt(texture)) .* speckle .* dop_shape;
y_rng = y_rng + clutter_rng;

Y = fft(y_rng, [], 1);
Y = Y + noise.sigma * (randn(Nsc,Nc) + 1j*randn(Nsc,Nc))/sqrt(2);
end

%% ========================================================================
%% 辅助函数（速度通道阈值可由 sys.SpeedSplit 注入）
%% ========================================================================

function groups = split_dets_by_speed(dets, nGroups, speedSplit)
% Split detections into speed groups and stamp det.cls (1..nGroups)
% dets: struct array with field .v (radial speed, m/s) OR .vr OR .fd (Hz) if available.
% nGroups: default 3  (slow/fast/vfast)
% speedSplit: optional thresholds [v1 v2] in m/s (abs speed); if empty -> auto-quantiles.

if nargin < 1 || isempty(dets)
    groups = cell(1,3);
    groups(:) = {struct([])};
    return;
end
if nargin < 2 || isempty(nGroups), nGroups = 3; end
if nargin < 3, speedSplit = []; end

% Extract speed estimate
v = nan(size(dets));
if isfield(dets,'v')
    v = [dets.v];
elseif isfield(dets,'vr')
    v = [dets.vr];
elseif isfield(dets,'spd')
    v = [dets.spd];
elseif isfield(dets,'fd')
    % If only doppler Hz, leave as-is (still works for grouping)
    v = [dets.fd];
end
v = abs(v(:));

% Determine thresholds
if isempty(speedSplit) || numel(speedSplit) < (nGroups-1)
    if all(isnan(v))
        thr = [];
    else
        vv = v(~isnan(v));
        if isempty(vv)
            thr = [];
        else
            % Quantile-based thresholds for robustness
            qs = linspace(0,1,nGroups+1);
            thr = quantile(vv, qs(2:end-1));
        end
    end
else
    thr = speedSplit(:).';
    thr = thr(1:(nGroups-1));
end

% Assign group id 1..nGroups
gid = ones(numel(dets),1);
if ~isempty(thr)
    for k=1:(nGroups-1)
        gid = gid + (v > thr(k));
    end
end
gid(isnan(v)) = 1; % unknown speed -> slow group

% Stamp and split
for i=1:numel(dets)
    dets(i).cls = gid(i);
end
groups = cell(1,nGroups);
for g=1:nGroups
    groups{g} = dets(gid==g);
end
end

function tracks_out = reorder_tracks_for_display(tracks_in, sys) %#ok<INUSD>
% 不依赖真值：用于可视化/面板展示的航迹排序
% 规则：LOCKED优先，其次按 hit_count / score 排序；保持结构体字段完全一致
if isempty(tracks_in)
    tracks_out = tracks_in;
    return;
end
N = numel(tracks_in);
locked = false(1,N);
hitc   = zeros(1,N);
score  = zeros(1,N);
age    = zeros(1,N);
for i=1:N
    if isfield(tracks_in(i),'status')
        locked(i) = strcmpi(tracks_in(i).status,'LOCKED');
    end
    hitc(i)  = getfield_safe(tracks_in(i),'hit_count',0);
    score(i) = getfield_safe(tracks_in(i),'score',0);
    age(i)   = getfield_safe(tracks_in(i),'age',0);
end
% sortrows: [-locked, -hitc, -score, -age]
M = [-double(locked(:)) -hitc(:) -score(:) -age(:)];
[~,ord] = sortrows(M);
tracks_out = tracks_in(ord);
end

function tracks_out = reorder_tracks_by_truth(tracks_in, targets) %#ok<INUSD>
% 保持函数名不变，但不使用targets真值重排（改为：按cls固定展示槽位）
% 规则：
%   slot1=SLOW, slot2=FAST, slot3=VFAST
% 每个槽位选择coh_gain_db最高的LOCKED，否则选TENTATIVE里质量最高（hits/score/coh）

tracks_out = repmat(struct( ...
    'id',nan,'pos',[],'vel',[],'speed',nan,'cls','', ...
    'status','SEARCHING','coh_gain_db',nan), 1, 3);

if isempty(tracks_in), return; end

want = {'SLOW','FAST','VFAST'};
for s=1:3
    cls = want{s};

    cand = [];
    for i=1:numel(tracks_in)
        if isfield(tracks_in(i),'cls') && strcmpi(tracks_in(i).cls, cls) && ~isempty(tracks_in(i).pos)
            cand = [cand, tracks_in(i)];
        end
    end
    if isempty(cand), continue; end

    % 先挑LOCKED
    lockMask = false(1,numel(cand));
    for i=1:numel(cand)
        lockMask(i) = isfield(cand(i),'status') && strcmpi(cand(i).status,'LOCKED');
    end

    pick = [];
    if any(lockMask)
        c2 = cand(lockMask);
        cg = arrayfun(@(x) getfield_safe(x,'coh_gain_db',-30), c2);
        [~,ii] = max(cg);
        pick = c2(ii);
    else
        % tentative按 hits/score/coh综合
        hits = arrayfun(@(x) getfield_safe(x,'hits',0), cand);
        sc   = arrayfun(@(x) getfield_safe(x,'score',3000), cand);
        cg   = arrayfun(@(x) getfield_safe(x,'coh_gain_db',-30), cand);
        q = 1.2*hits - 0.0010*sc + 0.05*cg;
        [~,ii] = max(q);
        pick = cand(ii);
    end

    % 关键修复：tracks_in 来自 mgr.tracks（字段很多），tracks_out 仅用于GUI显示（字段少）。
    % 不能直接结构体整体赋值，否则会触发“在不同结构体之间进行下标赋值”。
    if ~isempty(pick)
        gui = tracks_out(s);
        gui.id = getfield_safe(pick,'id',nan);
        if isfield(pick,'pos'), gui.pos = pick.pos; end
        if isfield(pick,'vel'), gui.vel = pick.vel; end
        if ~isempty(gui.vel) && all(isfinite(gui.vel))
            gui.speed = norm(gui.vel);
        else
            gui.speed = getfield_safe(pick,'speed',nan);
        end
        gui.cls = getfield_safe(pick,'cls',cls);
        gui.status = getfield_safe(pick,'status','SEARCHING');
        gui.coh_gain_db = getfield_safe(pick,'coh_gain_db',nan);
        tracks_out(s) = gui;
    end

end
end

function tracks = track_aided_coherent(tracks, rdBank, nodes, sys)
for k=1:numel(tracks)
    if isempty(tracks(k).pos) || ~isfield(tracks(k),'status') || ~strcmpi(tracks(k).status,'LOCKED')
        continue;
    end

    p = tracks(k).pos(:);
    v = tracks(k).vel(:);

    coh_sum = 0;
    incoh_sum = 0;
    used = 0;

    for rxIdx=1:3
        rx = nodes(rxIdx+2);
        for txIdx=1:2
            tx = nodes(txIdx);

            link_id = (rxIdx-1)*2 + txIdx;

            dsum = norm(p - tx.pos(:)) + norm(p - rx.pos(:));
            r_eq = 0.5 * dsum;
            cid_true = floor(r_eq / sys.R_unamb) + 1;
            if cid_true < 1 || cid_true > sys.Ncodes, continue; end

            r_fold = mod(r_eq, sys.R_unamb);
            tau = (2 * r_fold) / sys.c;

            u_tx = (p - tx.pos(:)); u_tx = u_tx/(norm(u_tx)+1e-12);
            u_rx = (p - rx.pos(:)); u_rx = u_rx/(norm(u_rx)+1e-12);
            v_meas_pred = dot(v, (u_tx+u_rx)/2);
            fd = 2*v_meas_pred/sys.lambda;

            rbin = round(r_fold/sys.r_res) + 1;
            rbin = max(1, min(sys.Nsc, rbin));
            [~,dbin] = min(abs(sys.fd_axis - fd));
            dbin = max(1, min(sys.Nc, dbin));

            RD = rdBank{link_id, cid_true};
            if isempty(RD), continue; end

            cellv = RD(rbin, dbin);
            cell_c = cellv * exp(1j*2*pi*sys.fc*tau);

            coh_sum   = coh_sum   + cell_c;
            incoh_sum = incoh_sum + abs(cellv)^2;
            used = used + 1;
        end
    end

    if used >= 3
        gain = 10*log10( (abs(coh_sum)^2 + 1e-12) / (incoh_sum + 1e-12) );
        if ~isfinite(gain), gain = -30; end
        tracks(k).coh_gain_db = gain;
    else
        tracks(k).coh_gain_db = -30;
    end
end
end

function display_tracks = get_display_tracks_from_mgr(mgr)
if ~isfield(mgr,'tracks') || isempty(mgr.tracks)
    display_tracks = struct('slot',{},'status',{},'pos',{},'vel',{},'cls',{},'coh_gain_db',{},'trkV',{});
    return;
end
tr = mgr.tracks;
N = numel(tr);
display_tracks = repmat(struct('slot','','status','SEARCHING','pos',[NaN NaN NaN], ...
    'vel',[NaN NaN NaN],'cls','','coh_gain_db',NaN,'trkV',NaN), 1, N);

for k=1:N
    display_tracks(k).slot = sprintf('TRK%d', k);
    if isfield(tr(k),'status') && ~isempty(tr(k).status), display_tracks(k).status = tr(k).status; end
    if isfield(tr(k),'pos') && ~isempty(tr(k).pos), display_tracks(k).pos = tr(k).pos; end
    if isfield(tr(k),'vel') && ~isempty(tr(k).vel), display_tracks(k).vel = tr(k).vel; end
    if isfield(tr(k),'cls') && ~isempty(tr(k).cls), display_tracks(k).cls = tr(k).cls; end
    if isfield(tr(k),'coh_gain_db') && ~isempty(tr(k).coh_gain_db), display_tracks(k).coh_gain_db = tr(k).coh_gain_db; end
    if isfield(tr(k),'vel') && ~isempty(tr(k).vel), display_tracks(k).trkV = norm(tr(k).vel); end
end
end
function gh = setup_gui(sys)
gh.fig = figure('Name','V22.1 PHANTOM SLAYER (修复增强版)','Color','w','Position',[50,50,1700,850]);
set(gh.fig,'Renderer','opengl'); drawnow;

gh.ax3  = subplot('Position',[0.03,0.33,0.46,0.63]); hold on; grid on; axis equal; view(3); box on; title('Battlefield 3D');
gh.axRD = subplot('Position',[0.53,0.33,0.44,0.63]); hold on; grid on; box on; title('Expanded RD'); xlabel('Range (km)'); ylabel('Velocity (m/s)');
xlim([0 300]); ylim([-sys.v_unamb sys.v_unamb]);

gh.axErr = subplot('Position',[0.53,0.05,0.44,0.20]); hold on; grid on; box on; title('Pos Error (m)'); ylim([0 1000]); xlabel('Time (s)'); ylabel('Error (m)');
gh.info = uicontrol('Style','text','Position',[50,20,700,240], 'FontName','Consolas','FontSize',11,'HorizontalAlignment','left', 'BackgroundColor',[0.95 0.95 0.95], 'String','');
end

function h = init_handles(gh, nodes, targets, sys)
h.ax3 = gh.ax3; h.axRD = gh.axRD; h.axErr = gh.axErr; h.info = gh.info;

axes(h.ax3);
h.nodes = gobjects(numel(nodes),1);
h.node_txt = gobjects(numel(nodes),1);
for i=1:numel(nodes)
    if strcmpi(nodes(i).type,'AIR')
        mk='^'; col='b';
    else
        mk='s'; col='k';
    end
    h.nodes(i)=plot3(nodes(i).pos(1),nodes(i).pos(2),nodes(i).pos(3),[col mk],'MarkerFaceColor',col);
    h.node_txt(i)=text(nodes(i).pos(1),nodes(i).pos(2),nodes(i).pos(3)+2000,nodes(i).name);
end

colors={'r','g','b'};
h.targets = gobjects(3,1); h.tgt_txt = gobjects(3,1); h.tracks = gobjects(3,1);
for i=1:3
    h.targets(i)=plot3(targets(i).pos(1),targets(i).pos(2),targets(i).pos(3),'^','Color',colors{i},'MarkerFaceColor',colors{i},'MarkerSize',9);
    h.tgt_txt(i)=text(targets(i).pos(1),targets(i).pos(2),targets(i).pos(3)+2000,targets(i).name,'Color',colors{i});
    h.tracks(i)=plot3(nan,nan,nan,'o','Color',colors{i},'LineWidth',2,'MarkerSize',8);
end

axes(h.axRD);
h.rd_img = imagesc(sys.x_km_sum, sys.v_axis, zeros(sys.Nc, sys.Nsc*sys.Ncodes));

axes(h.axErr);
h.err_lines(1)=plot(nan,nan,'r-');
h.err_lines(2)=plot(nan,nan,'g-');
h.err_lines(3)=plot(nan,nan,'b-');
end

function update_visuals(h, nodes, targets, tracks, rd_show_db, rt_err, sys)
for i=1:numel(nodes)
    set(h.nodes(i),'XData',nodes(i).pos(1),'YData',nodes(i).pos(2),'ZData',nodes(i).pos(3));
    set(h.node_txt(i),'Position',[nodes(i).pos(1),nodes(i).pos(2),nodes(i).pos(3)+2000]);
end

for i=1:numel(targets)
    set(h.targets(i),'XData',targets(i).pos(1),'YData',targets(i).pos(2),'ZData',targets(i).pos(3));
    set(h.tgt_txt(i),'Position',[targets(i).pos(1),targets(i).pos(2),targets(i).pos(3)+2000]);
end

for i=1:3
    if i<=numel(tracks) && ~isempty(tracks(i).pos) && all(isfinite(tracks(i).pos))
        p = tracks(i).pos;
        set(h.tracks(i),'XData',p(1),'YData',p(2),'ZData',p(3));
        if isfield(tracks(i),'status') && strcmpi(tracks(i).status,'LOCKED')
            set(h.tracks(i),'Marker','o','MarkerFaceColor',get(h.tracks(i),'Color'));
        else
            set(h.tracks(i),'Marker','o','MarkerFaceColor','none');
        end
    else
        set(h.tracks(i),'XData',nan,'YData',nan,'ZData',nan,'MarkerFaceColor','none');
    end
end

if ~isempty(rd_show_db)
    set(h.rd_img,'CData',rd_show_db);
    v = rd_show_db(:); v = v(isfinite(v));
    if ~isempty(v)
        hi = prctile(v, 99.7);
        lo = hi - 50;
        if (hi-lo) < 8, hi = lo+8; end
        caxis(h.axRD, [lo hi]);
    end
end

set(h.err_lines(1),'XData',rt_err.time,'YData',rt_err.e1);
set(h.err_lines(2),'XData',rt_err.time,'YData',rt_err.e2);
set(h.err_lines(3),'XData',rt_err.time,'YData',rt_err.e3);

allE = [rt_err.e1 rt_err.e2 rt_err.e3]; allE = allE(isfinite(allE));
if ~isempty(allE)
    yMax = 1000; %#ok<NASGU>
    ylim(h.axErr,[0 1000]);
end

xlim(h.axRD,[0 300]);
ylim(h.axRD,[-sys.v_unamb sys.v_unamb]);
end

function update_info_panel(hinfo, t, dets, solved, tracks, targets, rt_err, sys) %#ok<INUSD>
s = sprintf('Time: %.2f s | V22.1 PHANTOM SLAYER (Air-Only) | dt=0.5s\n', t);
s = [s, sprintf('det=%d | sol=%d | Nsc=%d | Ncodes=%d | R_unamb=%.1f km | RDx=[0,300]km\n', numel(dets), numel(solved), sys.Nsc, sys.Ncodes, sys.R_unamb/1e3)];
s = [s, sprintf('SpeedSplit: <230 / 230~400 / >400 (m/s)\n')];
s = [s, sprintf('%-10s | %-10s | %-10s | %-10s\n', 'Slot', 'Status', 'TrkV', 'CohGain(dB)')];

slotName = {'TRK1','TRK2','TRK3'}; % 显示用：按置信度排序后的前三条航迹

for k=1:3
    if k>numel(tracks)
        tr = struct('pos',[],'vel',[NaN;NaN;NaN],'status','SEARCHING','coh_gain_db',NaN,'id',NaN);
    else
        tr = tracks(k);
    end
    tag = slotName{k};

    if ~isempty(tr.pos) && isfield(tr,'status') && strcmpi(tr.status,'LOCKED')
        ev = norm(tr.vel);
        cg = getfield_safe(tr,'coh_gain_db',-30);
        tid = getfield_safe(tr,'id',nan);
        if isfinite(tid)
            tag = sprintf('%s(%d)', slotName{k}, round(tid));
        end
        s = [s, sprintf('%-10s | %-10s | %-10.0f | %-10.2f\n', tag, 'LOCKED', ev, cg)];
    else
        s = [s, sprintf('%-10s | %-10s | %-10s | %-10s\n', tag, 'SEARCHING', '-', '-')];
    end
end

% --- error summary (use recent-window RMSE + current error) ---
win = 10; % seconds
idxw = true(size(rt_err.time));
if isfield(rt_err,'time') && ~isempty(rt_err.time)
    idxw = rt_err.time >= max(0, t-win);
end
rmse = [sqrt(mean(rt_err.e1(idxw).^2,'omitnan')), ...
        sqrt(mean(rt_err.e2(idxw).^2,'omitnan')), ...
        sqrt(mean(rt_err.e3(idxw).^2,'omitnan'))];

cur = [nan nan nan];
if isfield(rt_err,'pos_err') && ~isempty(rt_err.pos_err)
    cur = rt_err.pos_err(end,:);
end

s = [s, sprintf('\nErrNow(m): J1=%.0f | J2=%.0f | J3=%.0f\n', cur(1), cur(2), cur(3))];
s = [s, sprintf('RMSE_%ds(m): J1=%.0f | J2=%.0f | J3=%.0f\n', win, rmse(1), rmse(2), rmse(3))];
set(hinfo, 'String', s);
end


function [sim_log, rt_err] = update_logs(sim_log, rt_err, t, targets, tracks)
% 注意：targets 在此仅用于“误差评估/后处理绘图”，不参与检测/关联/跟踪决策。
K = min(3, numel(targets));
N = numel(tracks);

% --- init ---
if isempty(rt_err)
    rt_err = struct();
end
if ~isfield(rt_err,'time'),  rt_err.time = []; end
if ~isfield(rt_err,'e1'),    rt_err.e1 = []; end
if ~isfield(rt_err,'e2'),    rt_err.e2 = []; end
if ~isfield(rt_err,'e3'),    rt_err.e3 = []; end
if ~isfield(rt_err,'coh1'),  rt_err.coh1 = []; end
if ~isfield(rt_err,'coh2'),  rt_err.coh2 = []; end
if ~isfield(rt_err,'coh3'),  rt_err.coh3 = []; end
if ~isfield(rt_err,'truth1'),rt_err.truth1 = nan(0,3); end
if ~isfield(rt_err,'truth2'),rt_err.truth2 = nan(0,3); end
if ~isfield(rt_err,'truth3'),rt_err.truth3 = nan(0,3); end
if ~isfield(rt_err,'est1'),  rt_err.est1 = nan(0,3); end
if ~isfield(rt_err,'est2'),  rt_err.est2 = nan(0,3); end
if ~isfield(rt_err,'est3'),  rt_err.est3 = nan(0,3); end

% --- append time (one sample per frame) ---
rt_err.time(end+1,1) = t;
idx = numel(rt_err.time);
rt_err.t = rt_err.time; % alias

% --- default NaN placeholders (keep lengths consistent) ---
rt_err.e1(idx,1) = NaN; rt_err.e2(idx,1) = NaN; rt_err.e3(idx,1) = NaN;
rt_err.coh1(idx,1) = NaN; rt_err.coh2(idx,1) = NaN; rt_err.coh3(idx,1) = NaN;
rt_err.truth1(idx,:) = nan(1,3); rt_err.truth2(idx,:) = nan(1,3); rt_err.truth3(idx,:) = nan(1,3);
rt_err.est1(idx,:)   = nan(1,3); rt_err.est2(idx,:)   = nan(1,3); rt_err.est3(idx,:)   = nan(1,3);

% --- nearest-neighbour pairing (evaluation only) ---
used = false(1,N);
for k = 1:K
    le.time = t;
    le.id = k;
    le.true_pos = targets(k).pos(:);
    le.true_vel = getfield_safe(targets(k),'vel',[NaN;NaN;NaN]);
    le.coh_gain_db = NaN;
    le.track_pos = [NaN;NaN;NaN];
    le.track_vel = [NaN;NaN;NaN];
    le.track_status = '';

    besti = 0; bestd = inf;
    for i=1:N
        if used(i), continue; end
        if ~isfield(tracks(i),'status') || ~strcmpi(tracks(i).status,'LOCKED'), continue; end
        if ~isfield(tracks(i),'pos') || isempty(tracks(i).pos) || any(~isfinite(tracks(i).pos(:))), continue; end
        d = norm(tracks(i).pos(:) - le.true_pos);
        if d < bestd
            bestd = d; besti = i;
        end
    end

    if besti > 0
        used(besti) = true;
        le.track_pos = tracks(besti).pos(:);
        le.track_vel = getfield_safe(tracks(besti),'vel',[NaN;NaN;NaN]);
        le.track_status = getfield_safe(tracks(besti),'status','');
        le.coh_gain_db = getfield_safe(tracks(besti),'coh_gain_db',NaN);
    else
        bestd = NaN;
    end

    sim_log = [sim_log, le]; %#ok<AGROW>

    % write per-target logs at same idx
    switch k
        case 1
            rt_err.truth1(idx,:) = le.true_pos(:).';
            rt_err.est1(idx,:)   = le.track_pos(:).';
            rt_err.e1(idx,1)     = bestd;
            rt_err.coh1(idx,1)   = le.coh_gain_db;
        case 2
            rt_err.truth2(idx,:) = le.true_pos(:).';
            rt_err.est2(idx,:)   = le.track_pos(:).';
            rt_err.e2(idx,1)     = bestd;
            rt_err.coh2(idx,1)   = le.coh_gain_db;
        case 3
            rt_err.truth3(idx,:) = le.true_pos(:).';
            rt_err.est3(idx,:)   = le.track_pos(:).';
            rt_err.e3(idx,1)     = bestd;
            rt_err.coh3(idx,1)   = le.coh_gain_db;
    end
end

% compatibility: pos_err
rt_err.pos_err = [rt_err.e1(:) rt_err.e2(:) rt_err.e3(:)];
end

function generate_final_report(logs)
% Optional offline report (kept for extension)
figure('Name','Analysis','Color','w','Position',[120,120,1100,520]);
subplot(1,2,1); hold on; grid on; axis equal; view(3); title('True vs Track Trajectories');
colors={'r','g','b'};
for k=1:3
    idx = [logs.id]==k; data = logs(idx); if isempty(data), continue; end
    tp = [data.true_pos]; tp=reshape(tp,3,[]);
    tr = [data.track_pos]; tr=reshape(tr,3,[]);
    plot3(tp(1,:),tp(2,:),tp(3,:),'-','Color',colors{k});
    plot3(tr(1,:),tr(2,:),tr(3,:),'x','Color',colors{k});
end

subplot(1,2,2); hold on; grid on; title('Position Error');
for k=1:3
    idx = [logs.id]==k; data = logs(idx); if isempty(data), continue; end
    tt = [data.time];
    errs = nan(1,numel(data));
    for j=1:numel(data)
        p = data(j).track_pos;
        if all(isnan(p)), errs(j)=nan; else, errs(j)=norm(p - data(j).true_pos); end
    end
    plot(tt, errs, 'Color', colors{k});
end
xlabel('Time (s)'); ylabel('Error (m)');
end

%% ========================================================================
%% 基础函数（保持不变）
%% ========================================================================

function alpha = os_cfar_alpha(Pfa, Ntrain, rank_frac)
alpha_ca = Ntrain * (Pfa^(-1/Ntrain) - 1);
alpha = alpha_ca * (0.85 + 0.5*(1-rank_frac));
end

function txSym = make_tx_symbols(sys)
txSym = cell(2, sys.Ncodes);
for cid=1:sys.Ncodes
    base1 = qpsk(sys.Nsc, 1);
    base2 = qpsk(sys.Nsc, 1);
    X1 = repmat(base1, 1, sys.Nc);
    X2 = repmat(base2, 1, sys.Nc);
    A = zeros(sys.Nsc, sys.Nc); B = zeros(sys.Nsc, sys.Nc);
    A(sys.sc_tx1,:) = X1(sys.sc_tx1,:);
    B(sys.sc_tx2,:) = X2(sys.sc_tx2,:);
    txSym{1,cid} = A;
    txSym{2,cid} = B;
end
end

function X = qpsk(N, M)
bits = randi([0,1], 2*N*M, 1);
b0 = bits(1:2:end); b1 = bits(2:2:end);
sym = (2*b0-1) + 1j*(2*b1-1);
sym = sym / sqrt(2);
X = reshape(sym, N, M);
end

function mgr = init_tracker_manager(priors, sys)
% Initialize tracker manager using priors (can be truth or any strong prior).
% Strong convention: slot/class mapping is fixed: 1=SLOW, 2=FAST, 3=VFAST.

mgr = struct();
mgr.time = 0;

% Track template (ensures consistent fields and avoids struct-assignment errors)
% make_track_struct signature is (id, slot)
    tmpl = make_track_struct(0, 0);
mgr.tracks = tmpl([]);  % empty struct array with template fields

if nargin < 1 || isempty(priors)
    return;
end

% Reorder priors into [SLOW, FAST, VFAST] if possible, but keep original values.
pri = priors; % keep original order (Target1/2/3)
	% strong slot↔class binding
	if isfield(sys,'TrackSlotClass') && numel(sys.TrackSlotClass)>=numel(pri)
		slot_cls = sys.TrackSlotClass;
	else
		slot_cls = {'SLOW','FAST','VFAST'};
	end

for k = 1:numel(pri)
    tr = tmpl;          % start from template => identical fields
    tr.id     = k;
    tr.name   = getfield_safe(pri(k), 'name', sprintf('TRK%d', k));
    tr.cls    = slot_cls{k};  % expected class (strong binding)
        tr.prior_cls = slot_cls{k};
    tr.slot   = k;       % strong slot mapping
    tr.tgt_id = getfield_safe(pri(k), 'id', k);

    % State: [x y z vx vy vz]'
    p0 = getfield_safe(pri(k), 'pos', [0 0 0]);
    v0 = getfield_safe(pri(k), 'vel', [0 0 0]);
    tr.x  = [p0(:); v0(:)];
    tr.P  = eye(6) * 5e6;

    % Make prior helpful but not "hard lock"
    tr.status = 'SEARCHING';
    tr.hits   = 0;
    tr.miss   = 0;
    tr.age    = 0;
    tr.last_meas_time = -inf;
    tr.last_meas_type = '';

    mgr.tracks(end+1) = tr; %#ok<AGROW>
end
end


function tr = make_track_struct(varargin)
% make_track_struct(id, slot) with optional args
if nargin >= 1, id = varargin{1}; else, id = 0; end
if nargin >= 2, slot = varargin{2}; else, slot = 0; end
% 3D CV (x,y,z,vx,vy,vz) track template with a STABLE field set.
% Keep this field set consistent everywhere to avoid struct-assignment errors.
if nargin<1 || isempty(id),   id   = 0; end
if nargin<2 || isempty(slot), slot = 0; end

tr = struct();
tr.id      = id;
	tr.tgt_id  = id;                   % 对应目标编号(用于伪量测关联)
tr.slot    = slot;                 % track slot index (1..N), 0 if unassigned
tr.name    = sprintf('TRK%d', id);
tr.cls     = 'UNK';                % 'SLOW'/'FAST'/'VFAST' or 'UNK'
tr.prior_cls = '';               % expected class from priors (optional)
                % 'SLOW'/'FAST'/'VFAST' or 'UNK'
tr.status  = 'SEARCHING';          % SEARCHING/TENTATIVE/LOCKED/LOST
tr.age     = 0;
tr.hits    = 0;
tr.miss    = 0;
tr.last_t  = -inf;
tr.last_meas_time = -inf;
tr.last_meas_type = '';

% state and covariance (CV model)
tr.x       = zeros(6,1);           % [x;y;z;vx;vy;vz]
tr.P       = diag([5e4,5e4,5e4, 5e2,5e2,5e2]).^2;

% convenience copies for display (kept updated in ensure_track_fields_fixed)
tr.pos     = zeros(1,3);
tr.vel     = zeros(1,3);
tr.speed   = 0;

% bookkeeping
tr.last_sol      = struct();       % last associated solution (struct)
tr.coh_gain_db   = nan;            % coherence gain estimate (dB)
tr.q             = 1.0;            % process-noise scaling (optional)

end

function cls = infer_cls_from_speed(speed_mps, sys_or_split)
% Infer class label from speed (m/s).
% sys_or_split can be:
%   - sys struct containing field SpeedSplit = [slow_fast, fast_vfast]
%   - numeric vector [slow_fast, fast_vfast]
if nargin < 2 || isempty(sys_or_split)
    split = [230 400];
elseif isstruct(sys_or_split)
    if isfield(sys_or_split,'SpeedSplit') && ~isempty(sys_or_split.SpeedSplit) && numel(sys_or_split.SpeedSplit) >= 2
        split = double(sys_or_split.SpeedSplit(1:2));
    else
        split = [230 400];
    end
else
    split = double(sys_or_split);
    if numel(split) < 2, split = [230 400]; else, split = split(1:2); end
end

spd = abs(double(speed_mps));
if spd > split(2)
    cls = 'VFAST';
elseif spd > split(1)
    cls = 'FAST';
else
    cls = 'SLOW';
end
end

function cls = infer_cls_from_prior(pr, sys)
v = getfield_safe(pr,'vel',[]);
if isempty(v), cls='SLOW'; return; end
spd = norm(v(:));
cls = infer_cls_from_speed(spd, sys);
end

function pri = reorder_priors_by_class_fixed(priors, sys)
% Reorder priors to match the fixed slot mapping: [SLOW, FAST, VFAST].
% If a class is missing, keep existing order for remaining ones.

pri = priors;
if isempty(priors), return; end

cls = cell(1,numel(priors));
for k=1:numel(priors)
    cls{k} = infer_cls_from_prior(priors(k), sys);
end

order = [];
for want = {"SLOW","FAST","VFAST"}
    idx = find(strcmpi(cls, want{1}), 1, 'first');
    if ~isempty(idx)
        order(end+1) = idx; %#ok<AGROW>
    end
end

% append any remaining priors
rest = setdiff(1:numel(priors), order, 'stable');
order = [order, rest];
pri = priors(order);
end
function q = infer_q_from_cls(cls)
% 过程噪声：越快越大（可按需要再调）
switch upper(string(cls))
    case "SLOW"
        q = 2.0;
    case "FAST"
        q = 5.0;
    otherwise
        q = 12.0;
end
end

function st = promote_status_fixed(st, hit_streak)
% 状态晋升逻辑（更容易进入 LOCKED）
if strcmpi(st,'LOCKED')
    st = 'LOCKED';
elseif hit_streak >= 3
    st = 'LOCKED';
elseif hit_streak >= 1
    st = 'TENTATIVE';
else
    st = 'SEARCHING';
end
end

function tracks = keep_best_3_tracks_fixed(tracks)
% 仅保留 slot=1..3 的轨迹；若无 slot 则取最优 3 条
if isempty(tracks), return; end

if isfield(tracks,'slot') && any(isfinite([tracks.slot]))
    % 只保留 slot=1..3；并确保每槽位最多一条
    out = repmat(make_track_struct(),1,0);
    for k=1:3
        idx = find([tracks.slot]==k);
        if isempty(idx)
            continue;
        end
        cand = tracks(idx);
        score = arrayfun(@(t) 20*getfield_safe(t,'miss_streak',0) - 2*getfield_safe(t,'hit_streak',0), cand);
        [~,ii] = min(score);
        out(end+1) = cand(ii); %#ok<AGROW>
    end
    tracks = out;
else
    % 没有 slot，就取 miss 最小 hit 最大的前三条
    score = arrayfun(@(t) 20*getfield_safe(t,'miss_streak',0) - 2*getfield_safe(t,'hit_streak',0), tracks);
    [~,ord] = sort(score,'ascend');
    tracks = tracks(ord(1:min(3,numel(ord))));
    for k=1:numel(tracks), tracks(k).slot = k; end
end
end

function pairs = assign_tracks_to_solutions_priors(tracks, sols, priors, sys)
% Legacy helper kept for compatibility; returns empty pairs (tracker uses solve_assignment_bruteforce_fixed).
pairs = zeros(numel(tracks),1);
end
function p = get_track_pos(tr)
% 兼容：track 可能只有 x(1:3)，也可能有 pos 字段
if isstruct(tr) && isfield(tr,'pos') && ~isempty(tr.pos) && isnumeric(tr.pos) && numel(tr.pos)>=3 && all(isfinite(tr.pos(:)))
    p = tr.pos(:);
elseif isstruct(tr) && isfield(tr,'x') && ~isempty(tr.x) && isnumeric(tr.x) && numel(tr.x)>=3 && all(isfinite(tr.x(1:3)))
    p = tr.x(1:3);
else
    p = [NaN;NaN;NaN];
end
end

function v = get_track_vel(tr)
% 兼容：track 可能只有 x(4:6)，也可能有 vel 字段
if isstruct(tr) && isfield(tr,'vel') && ~isempty(tr.vel) && isnumeric(tr.vel) && numel(tr.vel)>=3 && all(isfinite(tr.vel(:)))
    v = tr.vel(:);
elseif isstruct(tr) && isfield(tr,'x') && ~isempty(tr.x) && isnumeric(tr.x) && numel(tr.x)>=6 && all(isfinite(tr.x(4:6)))
    v = tr.x(4:6);
else
    v = [NaN;NaN;NaN];
end
end

function v = getfield_safe(s, f, d)
% Safe getter that works for numeric/logical as well as char/string/cell fields.
% If field is missing or empty, returns default d.
if isstruct(s) && isfield(s,f)
    x = s.(f);
    if isempty(x)
        v = d;
        return;
    end
    if isnumeric(x) || islogical(x)
        if all(isfinite(x(:)))
            v = x;
        else
            v = d;
        end
    else
        % Non-numeric (e.g., status strings, cell arrays, structs)
        v = x;
    end
else
    v = d;
end
end

function idx = cls_to_index_fixed(cls_tr, slot_i)
% Map class label (slow/fast/vfast) or numeric slot to 1..3 index.
% Strong constraint: slow->1, fast->2, vfast->3. Fallback to slot_i.
idx = [];
% If slot_i is valid numeric scalar, keep as fallback
if nargin < 2 || isempty(slot_i)
    slot_i = [];
end
if isnumeric(slot_i) && isscalar(slot_i) && isfinite(slot_i)
    idx = round(slot_i);
end

% Try parse cls_tr
if nargin >= 1 && ~isempty(cls_tr)
    if isstring(cls_tr), cls_tr = char(cls_tr); end
    if iscell(cls_tr) && ~isempty(cls_tr), cls_tr = cls_tr{1}; end
    if ischar(cls_tr)
        c = lower(strtrim(cls_tr));
        if contains(c,'slow'),  idx = 1; end
        if contains(c,'fast'),  idx = 2; end
        if contains(c,'vfast') || contains(c,'veryfast') || contains(c,'vf'), idx = 3; end
        % If it's like '1'/'2'/'3'
        if isempty(idx)
            v = str2double(c);
            if isfinite(v), idx = round(v); end
        end
    elseif isnumeric(cls_tr) && isscalar(cls_tr) && isfinite(cls_tr)
        idx = round(cls_tr);
    end
end

% Clamp
if isempty(idx) || ~isfinite(idx)
    idx = 1;
end
idx = max(1, min(3, idx));
end


function s2 = align_struct_to_template_fixed(s, tmpl)
% 对齐结构体字段：保证与 tmpl 字段集合与顺序一致。
% orderfields() 在字段不一致时会报错；这里用模板重建更稳。
if isempty(tmpl)
    s2 = s;
    return;
end
fn = fieldnames(tmpl);
s2 = struct();
for k = 1:numel(fn)
    f = fn{k};
    if isfield(s, f)
        s2.(f) = s.(f);
    else
        % 缺字段：用模板默认（若为空则用 []）
        try
            s2.(f) = tmpl.(f);
            if isempty(s2.(f)), s2.(f) = []; end
        catch
            s2.(f) = [];
        end
    end
end
end

function y = ternary(cond,a,b)
if cond, y=a; else, y=b; end
end

function q = q_from_cls_fixed(cls)
% Map class label to process-noise intensity for CV Kalman filter.
% The numeric value q scales the continuous white-acceleration model.
if isempty(cls)
    cls = 'FAST';
end
if strcmpi(cls,'SLOW')
    sa = 2.0;      % m/s^2
elseif strcmpi(cls,'FAST')
    sa = 12.0;     % m/s^2
else
    sa = 24.0;     % m/s^2 (VFAST)
end
q = sa^2;
end

function perf = efficiency_boundary_report(logs, nodes)
K = max([logs.id]);
colors = {'r','g','b'};

t_all = unique([logs.time]); %#ok<NASGU>
perf.rmse = nan(1,K);
perf.p90  = nan(1,K);
perf.pd   = nan(1,K);
perf.coh_med = nan(1,K);

cn = zeros(3,1);
for i=1:numel(nodes), cn = cn + nodes(i).pos(:); end
cn = cn/numel(nodes);

errs = cell(1,K);
rngs = cell(1,K);
cohg = cell(1,K);
for k=1:K
    idx = [logs.id]==k;
    dk = logs(idx);
    ek = nan(1,numel(dk));
    rk = nan(1,numel(dk));
    cgk = nan(1,numel(dk));
    for j=1:numel(dk)
        tp = dk(j).true_pos;
        tr = dk(j).track_pos;
        if all(isfinite(tr))
            ek(j)=norm(tr - tp);
        end
        rk(j)=norm(tp - cn);
        if isfield(dk(j),'coh_gain_db')
            cgk(j)=dk(j).coh_gain_db;
        end
    end
    errs{k}=ek; rngs{k}=rk/1e3; cohg{k}=cgk;

    ee = ek(isfinite(ek));
    if ~isempty(ee)
        perf.rmse(k) = sqrt(mean(ee.^2));
        perf.p90(k)  = prctile(ee,90);
        perf.pd(k)   = mean(ee < 1000);
    else
        perf.rmse(k)=NaN; perf.p90(k)=NaN; perf.pd(k)=0;
    end
    cg = cgk(isfinite(cgk));
    if ~isempty(cg), perf.coh_med(k)=median(cg); end
end

figure('Name','Efficiency Boundary','Color','w','Position',[120,120,1200,720]);

subplot(2,2,1); hold on; grid on; box on;
title('Empirical Pd vs Range (err<1km)');
xlabel('Range to node centroid (km)'); ylabel('Pd');
edges = 40:5:120;
cent = (edges(1:end-1)+edges(2:end))/2;
for k=1:K
    r = rngs{k}; e = errs{k};
    pd_bin = nan(1,numel(cent));
    for b=1:numel(cent)
        ii = r>=edges(b) & r<edges(b+1) & isfinite(e);
        if any(ii), pd_bin(b)=mean(e(ii)<1000); else, pd_bin(b)=NaN; end
    end
    plot(cent, pd_bin, 'Color', colors{k}, 'LineWidth', 1.8);
end
legend('Jet1','Jet2','Jet3','Location','southwest');

subplot(2,2,2); hold on; grid on; box on;
title('Position Error CDF');
xlabel('Error (m)'); ylabel('CDF');
for k=1:K
    ee = errs{k}; ee = ee(isfinite(ee));
    if isempty(ee), continue; end
    [f,x] = ecdf(ee);
    plot(x,f,'Color',colors{k},'LineWidth',1.8);
end
legend('Jet1','Jet2','Jet3','Location','southeast');

subplot(2,2,3); hold on; grid on; box on;
title('Position Error vs Time');
xlabel('Time (s)'); ylabel('Error (m)');
for k=1:K
    idx = [logs.id]==k;
    dk = logs(idx);
    tt = [dk.time];
    ee = errs{k};
    plot(tt, ee, '.', 'Color', colors{k});
end

subplot(2,2,4); hold on; grid on; box on;
title('Coherent Synthesis Advantage (proxy)');
xlabel('Range (km)'); ylabel('Effective gain (dB)');
for k=1:K
    r = rngs{k};
    cg = cohg{k};
    g_bin = nan(1,numel(cent));
    for b=1:numel(cent)
        ii = r>=edges(b) & r<edges(b+1) & isfinite(cg);
        if any(ii), g_bin(b)=median(cg(ii)); end
    end
    plot(cent, g_bin, 'Color', colors{k}, 'LineWidth', 1.8);
end
legend('Jet1','Jet2','Jet3','Location','southwest');

figure('Name','3D Detection Power Map','Color','w','Position',[160,160,1100,520]);
hold on; grid on; box on; view(3);
title('Detection Power (Pd surface over Range & Target)');
xlabel('Range (km)'); ylabel('Target'); zlabel('Pd(err<1km)');
[Yg,Xg] = meshgrid(1:K, cent);
Z = nan(size(Xg));
for k=1:K
    r = rngs{k}; e = errs{k};
    for b=1:numel(cent)
        ii = r>=edges(b) & r<edges(b+1) & isfinite(e);
        if any(ii), Z(b,k)=mean(e(ii)<1000); end
    end
end
surf(Xg,Yg,Z,'EdgeAlpha',0.25);
yticks(1:K); yticklabels({'Jet1','Jet2','Jet3'});
zlim([0 1]);
end


function assign = auction_min_cost(cost, max_cost)
% Auction algorithm for rectangular assignment (min cost)
% cost: nT x nS, inf means forbidden; max_cost used as gating fallback
% returns assign(nT,1) in 1..nS or 0 (unassigned)
if nargin<2 || isempty(max_cost), max_cost = inf; end
[nT,nS] = size(cost);
assign = zeros(nT,1);
if nT==0 || nS==0, return; end

% Replace inf with large number
big = max_cost;
if ~isfinite(big)
    finiteC = cost(isfinite(cost));
    if isempty(finiteC), return; end
    big = max(finiteC) + 1;
end
C = cost;
C(~isfinite(C)) = big;

% Convert to profit (maximize)
P = -C;
prices = zeros(1,nS);
owner  = zeros(1,nS); % which track owns item j

eps = 1e-3; % small epsilon for convergence
unassigned = 1:nT;
it_guard = 0;
it_max = 5000;
while ~isempty(unassigned) && it_guard < it_max
    it_guard = it_guard + 1;
    i = unassigned(1);
    % compute best and second best net values
    net = P(i,:) - prices;
    [v1,j1] = max(net);
    net(j1) = -inf;
    v2 = max(net);
    if ~isfinite(v2), v2 = v1 - 1; end
    bid = (v1 - v2) + eps;
    prices(j1) = prices(j1) + bid;

    prevOwner = owner(j1);
    owner(j1) = i;
    assign(i) = j1;
    if prevOwner~=0
        assign(prevOwner) = 0;
        unassigned = [unassigned(2:end) prevOwner]; %#ok<AGROW>
    else
        unassigned = unassigned(2:end);
    end
end

% 允许“禁配”：若该分配对应原始cost为inf/big，则置0
for i=1:nT
    j = assign(i);
    if j==0, continue; end
    if cost(i,j) > max_cost || ~isfinite(cost(i,j))
        assign(i) = 0;
    end
end
end

function v = getfield_def(s, f, def)
% Safe getfield with default (supports structs and empty).
v = def;
if nargin < 3, def = []; v = def; end
if isstruct(s) && isfield(s,f)
    try
        vv = s.(f);
        if ~isempty(vv)
            v = vv;
        end
    catch
    end
end
end

function cls = infer_cls_from_speed_fixed(speed_mps)
if isempty(speed_mps) || ~isfinite(speed_mps), speed_mps = 0; end
if speed_mps < 230
    cls = 'SLOW';
elseif speed_mps < 400
    cls = 'MED';
else
    cls = 'FAST';
end
end

function tr = make_track_slot_fixed(slot, pri, P0_pos, P0_vel)
% Create a track for a given slot using a prior (soft guidance).
% Fields MUST be consistent across the whole struct array.

if nargin < 3 || isempty(P0_pos), P0_pos = 2e6; end
if nargin < 4 || isempty(P0_vel), P0_vel = 2e4; end

slot = double(slot);
if isfield(pri,'pos') && ~isempty(pri.pos), pos0 = pri.pos(:); else, pos0 = zeros(3,1); end
if isfield(pri,'vel') && ~isempty(pri.vel), vel0 = pri.vel(:); else, vel0 = zeros(3,1); end

% base template
tr = track_template_fixed();
tr.id   = slot;
tr.slot = slot;

if isfield(pri,'name') && ~isempty(pri.name), tr.name = char(pri.name);
else, tr.name = sprintf('TRK%d',slot);
end

if isfield(pri,'cls') && ~isempty(pri.cls), tr.cls = char(pri.cls); end
tr.status = 'TENTATIVE';

% init state
tr.x(1:3) = pos0;
tr.x(4:6) = vel0;
tr.P      = diag([P0_pos P0_pos P0_pos P0_vel P0_vel P0_vel]);

tr.pos = tr.x(1:3);
tr.vel = tr.x(4:6);
tr.spd = norm(tr.vel);

% process noise scale
if isfield(pri,'q') && ~isempty(pri.q), tr.q = double(pri.q); end
end
function tr = ensure_track_fields_fixed(tr, slot)
if ~isfield(tr,'id') || isempty(tr.id), tr.id = slot; end
if ~isfield(tr,'slot') || isempty(tr.slot), tr.slot = slot; end
if ~isfield(tr,'name') || isempty(tr.name), tr.name = sprintf('TRK%d',slot); end
if ~isfield(tr,'cls'), tr.cls = ''; end
if ~isfield(tr,'status') || isempty(tr.status), tr.status = 'SEARCHING'; end
if ~isfield(tr,'hits') || isempty(tr.hits), tr.hits = 0; end
if ~isfield(tr,'hit_streak') || isempty(tr.hit_streak), tr.hit_streak = 0; end
if ~isfield(tr,'miss_streak') || isempty(tr.miss_streak), tr.miss_streak = 0; end
if ~isfield(tr,'coh_gain_db'), tr.coh_gain_db = NaN; end
if ~isfield(tr,'score'), tr.score = 0; end
if ~isfield(tr,'x') || isempty(tr.x)
    pos0 = [0;0;0]; vel0 = [0;0;0];
    if isfield(tr,'pos') && numel(tr.pos)==3, pos0 = tr.pos(:); end
    if isfield(tr,'vel') && numel(tr.vel)==3, vel0 = tr.vel(:); end
    tr.x = [pos0; vel0];
end
if ~isfield(tr,'P') || isempty(tr.P) || ~isequal(size(tr.P),[6 6])
    tr.P = diag([2e4^2 2e4^2 2e4^2 300^2 300^2 300^2]);
end
if ~isfield(tr,'pos') || isempty(tr.pos), tr.pos = tr.x(1:3).'; end
if ~isfield(tr,'vel') || isempty(tr.vel), tr.vel = tr.x(4:6).'; end
if ~isfield(tr,'last_sol'), tr.last_sol = struct(); end
end

function pairs = solve_assignment_bruteforce_fixed(cost, ok)
% For Nslot=3 only: return vector pairs(i)=assigned sol index or 0.
[Nt,Ns] = size(cost);
pairs = zeros(Nt,1);
if Nt==0 || Ns==0
    return;
end

% Build candidate lists
cand = cell(Nt,1);
for i=1:Nt
    cand{i} = find(ok(i,:));
    cand{i} = [cand{i} 0]; %#ok<AGROW> 0 means unassigned
end

best = inf;
bestPairs = zeros(Nt,1);

% brute force over small search space
for a = cand{1}
  for b = cand{min(2,Nt)}
    if Nt>=2 && a~=0 && b~=0 && a==b, continue; end
    for c = cand{min(3,Nt)}
        if Nt>=3
            if a~=0 && c~=0 && a==c, continue; end
            if b~=0 && c~=0 && b==c, continue; end
        end
        p = [a;b;c];
        p = p(1:Nt);
        val = 0;
        for i=1:Nt
            if p(i)==0, val = val + 1.2e9; % prefer assignment when possible
            else,        val = val + cost(i,p(i));
            end
        end
        if val < best
            best = val;
            bestPairs = p;
        end
    end
  end
end

pairs = bestPairs;
end

function tr = track_hit_step_fixed(tr, lock_hits)
tr.hits = tr.hits + 1;
tr.hit_streak = tr.hit_streak + 1;
tr.miss_streak = 0;
if tr.hit_streak >= lock_hits
    tr.status = 'LOCKED';
else
    tr.status = 'TENTATIVE';
end
end

function tr = track_miss_step_fixed(tr, drop_miss)
tr.miss_streak = tr.miss_streak + 1;
tr.hit_streak = max(0, tr.hit_streak - 1);
if tr.miss_streak >= drop_miss
    tr.status = 'SEARCHING';
end
end



function sic_targets = sols_to_targets(sols)
% Convert solution structs to target-like structs for SIC cancellation.
sic_targets = struct('id',{},'name',{},'pos',{},'vel',{},'vel0',{},'rcs',{});
if isempty(sols), return; end
for i = 1:numel(sols)
    p = getfield_safe(sols(i),'pos',[0 0 0]); p = p(:).';
    v = getfield_safe(sols(i),'vel',[0 0 0]); v = v(:).';
    sic_targets(i).id   = i;
    sic_targets(i).name = sprintf('SIC%d',i);
    sic_targets(i).pos  = p;
    sic_targets(i).vel  = v;
    sic_targets(i).vel0 = v;
    sic_targets(i).rcs  = getfield_safe(sols(i),'rcs', 1.0);
end
end


% === Track struct utilities (field-consistency) ===
function tmpl = track_template_fixed()
% Canonical track struct template. Keep this in sync with tracker usage.
tmpl = struct();
tmpl.id          = 0;
tmpl.slot        = 0;
tmpl.name        = '';
tmpl.cls         = '';
tmpl.status      = 'TENTATIVE';
tmpl.age         = 0;
tmpl.hits        = 0;
tmpl.miss        = 0;
tmpl.hit_streak  = 0;
tmpl.miss_streak = 0;
tmpl.x           = zeros(6,1);
tmpl.P           = eye(6);
tmpl.q           = 1;
tmpl.coh_gain_db = 0;
tmpl.score       = 0;
tmpl.last_t      = -inf;
tmpl.pos         = zeros(3,1);
tmpl.vel         = zeros(3,1);
tmpl.spd         = 0;
tmpl.last_sol    = struct();
end

function tracks = normalize_tracks_fixed(tracks, tmpl)
% Normalize track struct array to have EXACTLY the same fields as template `tmpl`.
% This avoids `orderfields` errors when tracks carry extra/debug fields.
%
% Policy:
%   - Missing fields are filled from template defaults.
%   - Extra fields are dropped (keeps tracker stable and predictable).
%   - Field order follows template.
%
% Input:
%   tracks : struct (scalar or array)
%   tmpl   : template struct (same fields desired)
%
% Output:
%   tracks : struct array with uniform fields/order
if nargin < 2 || isempty(tmpl)
    tmpl = track_template_fixed();
end

if isempty(tracks)
    tracks = repmat(tmpl, 0, 1);
    return;
end

fn = fieldnames(tmpl);

% Preallocate output with template (ensures correct fields/order)
tracks_out = repmat(tmpl, size(tracks));

for k = 1:numel(tracks)
    for i = 1:numel(fn)
        f = fn{i};
        if isfield(tracks(k), f)
            try
                v = tracks(k).(f);
                if ~isempty(v)
                    tracks_out(k).(f) = v;
                else
                    tracks_out(k).(f) = tmpl.(f);
                end
            catch
                tracks_out(k).(f) = tmpl.(f);
            end
        else
            tracks_out(k).(f) = tmpl.(f);
        end
    end

    % Keep derived fields consistent with state vector if present
    if isfield(tracks_out(k), 'x') && isnumeric(tracks_out(k).x) && numel(tracks_out(k).x) >= 6
        tracks_out(k).pos = tracks_out(k).x(1:3);
        tracks_out(k).vel = tracks_out(k).x(4:6);
        tracks_out(k).spd = norm(tracks_out(k).vel);
    end
end

tracks = tracks_out;
end
% Helper: invert Albersheim to get Pd for given SNR (numerical)
    function Pd = pd_from_snr_db(snr_db, Pfa, Np)
        % Invert Albersheim (approx) to get Pd from SNR for given Pfa and Np
        Pd_grid = linspace(0.05,0.999,400);
        snr_req = arrayfun(@(pd) albersheim_snr_db_fixed(Pfa,pd,Np), Pd_grid);
        % Robustify against complex/NaN/unsorted issues (avoid: "输入坐标必须为实数")
        snr_req = real(snr_req(:));
        Pd_grid = Pd_grid(:);
        m = isfinite(snr_req) & isfinite(Pd_grid);
        snr_req = snr_req(m);
        Pd_grid = Pd_grid(m);
        [snr_req, ord] = sort(snr_req);
        Pd_grid = Pd_grid(ord);
        [snr_req_u, iu] = unique(snr_req, 'stable');
        Pd_grid_u = Pd_grid(iu);
        Pd = interp1(snr_req_u, Pd_grid_u, real(snr_db), 'linear', 'extrap');
        Pd = real(Pd);
        Pd = min(max(Pd,0.001),0.999);
    end

function plot_advanced_perf_boundaries_1(sys, sim_log, rt_err)
% 高级效能边界分析（增强版，仅保留信号级CRLB与实际误差对比）
% 包含：
%   - 检测概率 vs 距离（理论与仿真统计）
%   - 信号级 CRLB 与实际误差（分目标，线型区分）
%   - 多链路相干合成增益
%   - 速度分区间检测概率（体现三通道并行检测优势）
%
% 注意：实际误差为卡尔曼滤波后的结果，可能略低于瞬时CRLB（因时间积累），
%       但仍可直观反映系统跟踪性能接近理论下界。

% ---------- 1. 参数提取与默认值 ----------
c       = sys.c;
fc      = sys.fc;
lambda  = c/fc;
B       = sys.B;
Nc      = sys.Nc;
PRI     = sys.PRI;
Tc      = Nc * PRI;                    % 相干积累时间
SNR0_dB = getfield_safe(sys, 'SNR0_dB', 18);
R0_km   = getfield_safe(sys, 'R0_km', 80);
R0_m    = R0_km * 1e3;
Pfa     = getfield_safe(sys, 'Pfa', 1e-6);
RCS_list = [25, 5, 0.5];                % Jet1, Jet2, Jet3
target_names = {'Jet1','Jet2','Jet3'};
colors = {'r','g','b'};

% 节点位置（近似固定，忽略运动）
nodes = make_nodes();
tx_pos = [nodes(1).pos; nodes(2).pos];
rx_pos = [nodes(3).pos; nodes(4).pos; nodes(5).pos];
L = size(tx_pos,1) * size(rx_pos,1);    % 总链路数 = 6

% ---------- 2. 检测概率曲线（理论 vs 仿真） ----------
% 理论计算：对距离网格求Pd
dist_grid = linspace(30, 250, 100) * 1e3;   % 30~250 km
Pd_theory_single = zeros(3, length(dist_grid));
Pd_theory_multi  = zeros(3, length(dist_grid));

for k = 1:3
    RCS = RCS_list(k);
    for j = 1:length(dist_grid)
        R = dist_grid(j);
        SNR_lin = 10^( (SNR0_dB + 10*log10(RCS) - 40*log10(R/R0_m)) / 10 );
        Pd_single = pd_swerling0(SNR_lin, Pfa);
        Pd_theory_single(k,j) = Pd_single;
        Pd_theory_multi(k,j) = 1 - (1-Pd_single)^L;
    end
end

% 从仿真统计Pd
if ~isempty(sim_log)
    dist_all = []; detect_all = []; id_all = [];
    for entry = sim_log
        if isempty(entry.true_pos), continue; end
        R = norm(entry.true_pos);
        dist_all(end+1) = R/1e3;
        detect_all(end+1) = isfield(entry,'track_pos') && ~isempty(entry.track_pos) && ...
                            norm(entry.track_pos - entry.true_pos) < 1000;
        id_all(end+1) = entry.id;
    end
    edges = 30:20:250; centers = (edges(1:end-1)+edges(2:end))/2;
    Pd_sim = nan(3, length(centers));
    for k = 1:3
        for b = 1:length(centers)
            idx = id_all==k & dist_all>=edges(b) & dist_all<edges(b+1);
            if sum(idx) >= 5
                Pd_sim(k,b) = mean(detect_all(idx));
            end
        end
    end
end

% 绘图：Pd vs 距离
figure('Name','检测概率曲线','Color','w','Position',[100,100,800,600]);
for k = 1:3
    subplot(1,3,k); hold on; grid on;
    plot(dist_grid/1e3, Pd_theory_single(k,:), '--', 'Color', colors{k}, 'LineWidth',1.5);
    plot(dist_grid/1e3, Pd_theory_multi(k,:), '-', 'Color', colors{k}, 'LineWidth',2);
    if exist('Pd_sim','var')
        plot(centers, Pd_sim(k,:), 'o', 'Color', colors{k}, 'MarkerFaceColor', colors{k});
    end
    xlabel('距离 (km)'); ylabel('P_d'); ylim([0 1]);
    title([target_names{k}, ' (RCS=',num2str(RCS_list(k)),' m^2)']);
    legend('单链路理论','多链路理论','仿真统计','Location','southwest');
end
sgtitle('检测概率：单链路 vs 多链路联合');

% ---------- 3. 跟踪误差CRLB与仿真对比（分目标，仅信号级CRLB与实际误差） ----------
if ~isempty(sim_log) && ~isempty(rt_err)
    % 按目标提取数据
    t_target = cell(1,3);
    crlb_sig_target = cell(1,3);   % 信号级CRLB（联合FIM，距离+速度）
    pos_err_target = cell(1,3);     % 实际位置误差
    
    for target_id = 1:3
        idx_target = [sim_log.id] == target_id;
        sim_tgt = sim_log(idx_target);
        if isempty(sim_tgt), continue; end
        Nt = numel(sim_tgt);
        t_target{target_id} = [sim_tgt.time];
        crlb_sig = nan(Nt,1);
        pos_err = nan(Nt,1);
        
        for i = 1:Nt
            entry = sim_tgt(i);
            xt = entry.true_pos(:)';
            vt = entry.true_vel(:)';
            RCS = RCS_list(target_id);
            
            % 初始化联合FIM（信号级，使用距离+速度）
            J_sig = zeros(6);
            
            link_idx = 0;
            for tx_idx = 1:2
                for rx_idx = 1:3
                    link_idx = link_idx + 1;
                    tx = tx_pos(tx_idx,:);
                    rx = rx_pos(rx_idx,:);
                    
                    R_tx = norm(xt - tx);
                    R_rx = norm(xt - rx);
                    if R_tx < 1 || R_rx < 1, continue; end
                    
                    SNR_lin = 10^( (SNR0_dB + 10*log10(RCS) - 40*log10(sqrt(R_tx*R_rx)/R0_m)) / 10 );
                    SNR_lin = max(SNR_lin, 1e-6);
                    
                    sigma_R = c/(2*B*sqrt(2*SNR_lin));   % 单程测距误差
                    sigma_r = sigma_R / sqrt(2);        % 双基距离和误差（方差减半）
                    sigma_v = lambda/(2*Tc*sqrt(SNR_lin)); % 双基径向速度误差
                    
                    [H_full, ~] = bistatic_jacobian(xt, vt, tx, rx); % 2x6雅可比
                    R_full = diag([sigma_r^2, sigma_v^2]);
                    J_sig = J_sig + H_full' * (R_full \ H_full);
                end
            end
            
            % 信号级CRLB
            J_sig_reg = J_sig + 1e-9 * eye(6);
            if cond(J_sig_reg) > 1e12
                P_sig = pinv(J_sig_reg);
            else
                P_sig = inv(J_sig_reg);
            end
            crlb_sig(i) = sqrt(trace(P_sig(1:3,1:3))); % 位置RMSE下界
            
            % 实际位置误差（从rt_err提取）
            tidx = find(abs(rt_err.time - entry.time) < 0.01, 1);
            if ~isempty(tidx)
                switch target_id
                    case 1, pos_err(i) = rt_err.e1(tidx);
                    case 2, pos_err(i) = rt_err.e2(tidx);
                    case 3, pos_err(i) = rt_err.e3(tidx);
                end
            end
        end
        
        crlb_sig_target{target_id} = crlb_sig;
        pos_err_target{target_id} = pos_err;
    end
    
    % 绘图：仅信号级CRLB与实际误差
    figure('Name','跟踪误差下界（分目标）','Color','w','Position',[400,400,1000,600]);
    hold on; grid on;
    
    for k = 1:3
        if isempty(t_target{k}), continue; end
        % 信号级CRLB（实线）
        plot(t_target{k}, crlb_sig_target{k}, '-', 'Color', colors{k}, 'LineWidth',2, ...
            'DisplayName',[target_names{k} ' 信号级CRLB']);
        % 实际误差（点线）
        plot(t_target{k}, pos_err_target{k}, ':', 'Color', colors{k}, 'LineWidth',1.5, ...
            'DisplayName',[target_names{k} ' 实际误差']);
    end
    
    xlabel('时间 (s)'); ylabel('位置RMSE (m)');
    title('信号级CRLB与实际误差（分目标）');
    legend('Location','best');
    ylim([0 1000]);  % 可根据实际数据调整
    hold off;
end

% ---------- 4. 相干增益曲线 ----------
if ~isempty(rt_err) && isfield(rt_err,'coh1')
    t = rt_err.time(:);
    coh = [rt_err.coh1(:), rt_err.coh2(:), rt_err.coh3(:)];
    figure('Name','相干合成增益','Color','w','Position',[450,450,800,400]);
    hold on; grid on;
    plot(t, 10*log10(L)*ones(size(t)), 'k--', 'LineWidth',2, 'DisplayName','理论最大增益');
    for k = 1:3
        plot(t, coh(:,k), 'Color', colors{k}, 'LineWidth',1.2, 'DisplayName',target_names{k});
    end
    xlabel('时间 (s)'); ylabel('相干增益 (dB)');
    title('多链路相干合成增益');
    legend('Location','best');
end

% ---------- 5. 速度分区间检测概率（体现三通道并行检测优势） ----------
if ~isempty(sim_log) && ~isempty(rt_err)
    v_edges = [0, 200, 400, 600];
    v_centers = (v_edges(1:end-1) + v_edges(2:end)) / 2;
    n_bins = length(v_centers);
    
    detect_count = zeros(3, n_bins);
    total_count = zeros(3, n_bins);
    
    for entry = sim_log
        if ~isfield(entry,'true_vel') || isempty(entry.true_vel), continue; end
        target_id = entry.id;
        if target_id < 1 || target_id > 3, continue; end
        v = norm(entry.true_vel);
        bin_idx = find(v >= v_edges(1:end-1) & v < v_edges(2:end), 1);
        if isempty(bin_idx), continue; end
        total_count(target_id, bin_idx) = total_count(target_id, bin_idx) + 1;
        
        tidx = find(abs(rt_err.time - entry.time) < 0.01, 1);
        if ~isempty(tidx)
            switch target_id
                case 1, err = rt_err.e1(tidx);
                case 2, err = rt_err.e2(tidx);
                case 3, err = rt_err.e3(tidx);
            end
            if ~isnan(err) && err < 1000
                detect_count(target_id, bin_idx) = detect_count(target_id, bin_idx) + 1;
            end
        end
    end
    
    detect_prob = detect_count ./ max(total_count, 1);
    detect_prob(total_count < 5) = NaN;
    
    figure('Name','速度分区间检测概率','Color','w','Position',[500,500,800,400]);
    bar_width = 0.25;
    x = 1:n_bins;
    for k = 1:3
        offset = (k-2) * bar_width;
        bar(x + offset, detect_prob(k,:), bar_width, 'FaceColor', colors{k}, ...
            'DisplayName', target_names{k});
        hold on;
    end
    hold off;
    grid on;
    xlabel('速度区间 (m/s)');
    ylabel('检测概率 P_d');
    title('三通道并行检测效果：不同速度区间检测概率');
    xticks(x);
    xticklabels({'0-200','200-400','400-600'});
    ylim([0 1]);
    legend('Location','best');
    
    % 添加样本数量标注
    for k = 1:3
        for b = 1:n_bins
            if total_count(k,b) >= 5
                text(x(b) + (k-2)*bar_width, detect_prob(k,b) + 0.02, ...
                    sprintf('n=%d', total_count(k,b)), ...
                    'HorizontalAlignment','center','FontSize',8);
            end
        end
    end
end

% ---------- 6. 辅助函数 ----------
    function Pd = pd_swerling0(SNR_lin, Pfa)
        x_th = chi2inv(1-Pfa, 2);
        Pd = 1 - ncx2cdf(x_th, 2, 2*SNR_lin);
    end

    function [H, rv] = bistatic_jacobian(x, v, tx, rx)
        p = x(:); v = v(:);
        d_tx = p - tx(:); R_tx = norm(d_tx);
        d_rx = p - rx(:); R_rx = norm(d_rx);
        u_tx = d_tx / R_tx;
        u_rx = d_rx / R_rx;
        
        % 距离对位置的导数
        dr_dp = 0.5 * (u_tx' + u_rx');   % 1x3
        dr_dv = zeros(1,3);
        
        % 径向速度对位置的导数（考虑u_tx,u_rx变化）
        du_tx_dp = (eye(3) - u_tx*u_tx') / R_tx;
        du_rx_dp = (eye(3) - u_rx*u_rx') / R_rx;
        dv_dp = 0.5 * (v'*du_tx_dp + v'*du_rx_dp); % 1x3
        % 径向速度对速度的导数
        dv_dv = 0.5 * (u_tx' + u_rx');   % 1x3
        
        H = [dr_dp, dr_dv; dv_dp, dv_dv];
        rv = [0.5*(R_tx+R_rx); 0.5*(u_tx'*v + u_rx'*v)]; % 测量预测
    end
end
function generate_all_performance_plots(sys, sim_log, rt_err, nodes, targets, all_dets_history, rdBank)
% 生成全套效能分析图（中文图名）- 最终稳定版
% 输入：
%   sys     - 系统参数结构体
%   sim_log - 仿真日志（每时刻每个目标的信息）
%   rt_err  - 实时误差记录
%   nodes   - 节点位置
%   targets - 目标真值
%   all_dets_history - 所有时刻的检测点历史（cell数组）
%   rdBank  - 复数RD缓存（6×sys.Ncodes cell，每个元素为 Nsc×Nc 复数矩阵）

% ---------- 1. 参数提取 ----------
c       = sys.c;
fc      = sys.fc;
lambda  = c/fc;
B       = sys.B;
Nc      = sys.Nc;
PRI     = sys.PRI;
Tc      = Nc * PRI;
SNR0_dB = getfield_safe(sys, 'SNR0_dB', 18);
R0_km   = getfield_safe(sys, 'R0_km', 80);
R0_m    = R0_km * 1e3;
Pfa     = getfield_safe(sys, 'Pfa', 1e-6);
RCS_list = [25, 5, 0.5];
target_names = {'Jet1','Jet2','Jet3'};
colors = {'r','g','b'};

tx_pos = [nodes(1).pos; nodes(2).pos];
rx_pos = [nodes(3).pos; nodes(4).pos; nodes(5).pos];
L = 6;

% ---------- 2. 辅助函数 ----------
    function SNR_dB = compute_SNR(R_m, RCS)
        SNR_dB = SNR0_dB - 40*log10(R_m / R0_m) + 10*log10(RCS);
    end
    function Pd = pd_swerling0(SNR_lin, Pfa)
        x_th = chi2inv(1-Pfa, 2);
        Pd = 1 - ncx2cdf(x_th, 2, 2*SNR_lin);
    end
    function [H, rv] = bistatic_jacobian(x, v, tx, rx)
        p = x(:); v = v(:);
        d_tx = p - tx(:); R_tx = norm(d_tx);
        d_rx = p - rx(:); R_rx = norm(d_rx);
        u_tx = d_tx / R_tx;
        u_rx = d_rx / R_rx;
        dr_dp = 0.5 * (u_tx' + u_rx');
        dr_dv = zeros(1,3);
        du_tx_dp = (eye(3) - u_tx*u_tx') / R_tx;
        du_rx_dp = (eye(3) - u_rx*u_rx') / R_rx;
        dv_dp = 0.5 * (v'*du_tx_dp + v'*du_rx_dp);
        dv_dv = 0.5 * (u_tx' + u_rx');
        H = [dr_dp, dr_dv; dv_dp, dv_dv];
        rv = [0.5*(R_tx+R_rx); 0.5*(u_tx'*v + u_rx'*v)];
    end
    function crlb = compute_crlb_signal(xt, vt, RCS, node_set)
        % 信号级CRLB，返回位置RMSE下界（米）
        try
            J = zeros(6);
            valid_links = 0;
            for idx = node_set
                if idx < 1 || idx > 6, continue; end
                [i,j] = ind2sub([2,3], idx);
                if i<1 || i>2 || j<1 || j>3, continue; end
                tx = tx_pos(i,:);
                rx = rx_pos(j,:);
                R_tx = norm(xt - tx);
                R_rx = norm(xt - rx);
                % 距离总是正数，无需过滤
                SNR_lin = 10^( (SNR0_dB + 10*log10(RCS) - 40*log10(sqrt(R_tx*R_rx)/R0_m)) / 10 );
                SNR_lin = max(SNR_lin, 1e-12);
                sigma_r = c/(2*B*sqrt(2*SNR_lin)) / sqrt(2);
                sigma_v = lambda/(2*Tc*sqrt(SNR_lin));
                [H, ~] = bistatic_jacobian(xt, vt, tx, rx);
                R = diag([sigma_r^2, sigma_v^2]);
                J = J + H' * (R \ H);
                valid_links = valid_links + 1;
            end
            if valid_links < 3
                crlb = NaN; % 链路不足时返回NaN，绘图时跳过
                return;
            end
            % 自适应正则化
            reg_scale = max(max(diag(J)), 1e-6);
            J_reg = J + 1e-3 * eye(6) * reg_scale;
            P = pinv(J_reg);
            crlb = sqrt(trace(P(1:3,1:3)));
            if ~isfinite(crlb)
                crlb = NaN;
            end
        catch
            crlb = NaN;
        end
    end
    function crlb_info = compute_crlb_info_range(xt, RCS, node_set)
        % 信息级CRLB（仅距离量测）
        try
            J_pos = zeros(3);
            valid_links = 0;
            for idx = node_set
                if idx < 1 || idx > 6, continue; end
                [i,j] = ind2sub([2,3], idx);
                if i<1 || i>2 || j<1 || j>3, continue; end
                tx = tx_pos(i,:);
                rx = rx_pos(j,:);
                R_tx = norm(xt - tx);
                R_rx = norm(xt - rx);
                SNR_lin = 10^( (SNR0_dB + 10*log10(RCS) - 40*log10(sqrt(R_tx*R_rx)/R0_m)) / 10 );
                SNR_lin = max(SNR_lin, 1e-12);
                sigma_r = c/(2*B*sqrt(2*SNR_lin)) / sqrt(2);
                u_tx = (xt - tx)/R_tx;
                u_rx = (xt - rx)/R_rx;
                dr_dp = 0.5 * (u_tx + u_rx);
                J_pos = J_pos + (dr_dp' * dr_dp) / sigma_r^2;
                valid_links = valid_links + 1;
            end
            if valid_links < 3
                crlb_info = NaN;
                return;
            end
            reg_scale = max(max(diag(J_pos)), 1e-6);
            J_pos_reg = J_pos + 1e-3*eye(3) * reg_scale;
            P = pinv(J_pos_reg);
            crlb_info = sqrt(trace(P));
            if ~isfinite(crlb_info)
                crlb_info = NaN;
            end
        catch
            crlb_info = NaN;
        end
    end
    function snr_db = albersheim_snr_db_fixed(Pfa, Pd, N)
        Pd = min(max(Pd, 1e-6), 1-1e-6);
        Pfa = min(max(Pfa, 1e-12), 1-1e-12);
        N = max(1, round(N));
        A = log(0.62./Pfa);
        B = log(Pd./(1-Pd));
        snr_lin = (A + 0.12.*A.*B + 1.7.*B) ./ N;
        snr_lin = real(max(snr_lin, 1e-12));
        snr_db = 10*log10(snr_lin);
    end
    function RD = generate_ideal_rd(tx, rx, xt, vt, sys)
        % 生成理想单目标RD图，返回尺寸 Nsc × Nc
        Nsc = sys.Nsc; Nc = sys.Nc;
        X = ones(Nsc, Nc);  % 简化波形
        d_tx = norm(xt - tx);
        d_rx = norm(xt - rx);
        r_eq = 0.5 * (d_tx + d_rx);
        r_fold = mod(r_eq, sys.R_unamb);
        tau = 2 * r_fold / sys.c;
        u_tx = (xt - tx)/d_tx;
        u_rx = (xt - rx)/d_rx;
        v_rad = 0.5 * (dot(vt, u_tx) + dot(vt, u_rx));
        fd = 2 * v_rad / sys.lambda;
        f_k = (0:Nsc-1)' * sys.df;
        carrier = exp(-1j*2*pi*sys.fc*tau);
        ph_k = exp(-1j*2*pi*f_k*tau);
        ph_d = exp(1j*2*pi*fd*sys.PRI*(0:Nc-1));
        amp = 1;
        Y = amp * carrier * (ph_k * ph_d) .* X; % Y: Nsc × Nc
        Z = Y .* conj(X);                       % Z: Nsc × Nc
        z_mf = ifft(Z, [], 1);                   % z_mf: Nsc × Nc
        w = hann(Nc).';
        RD = fftshift(fft(z_mf .* w, [], 2), 2); % 结果仍为 Nsc × Nc
    end

% ---------- 3. 准备数据 ----------
if isempty(sim_log)
    warning('sim_log为空，无法生成后续图形。');
    return;
end
t_all = [sim_log.time];
xt_all = zeros(3, numel(sim_log));
vt_all = zeros(3, numel(sim_log));
id_all = [sim_log.id];
for k = 1:numel(sim_log)
    xt_all(:,k) = sim_log(k).true_pos(:);
    vt_all(:,k) = sim_log(k).true_vel(:);
end
t_rt = rt_err.time;
e1 = rt_err.e1(:); e2 = rt_err.e2(:); e3 = rt_err.e3(:);

% ---------- 定义公共距离网格 ----------
dist_grid = linspace(30, 250, 100) * 1e3;   % 30~250 km

% ---------- 图1：检测概率 vs 距离（理论+仿真） ----------
try
    Pd_theory = zeros(3, length(dist_grid));
    Pd_theory_single = zeros(3, length(dist_grid));
    Pd_theory_multi  = zeros(3, length(dist_grid));

    for k = 1:3
        RCS = RCS_list(k);
        for j = 1:length(dist_grid)
            R = dist_grid(j);
            SNR_lin = 10^( (SNR0_dB + 10*log10(RCS) - 40*log10(R/R0_m)) / 10 );
            Pd_single = pd_swerling0(SNR_lin, Pfa);
            Pd_theory_single(k,j) = Pd_single;
            Pd_theory_multi(k,j) = 1 - (1-Pd_single)^L;
        end
    end

    if ~isempty(sim_log)
        dist_all = []; detect_all = []; id_all_stat = [];
        for entry = sim_log
            if isempty(entry.true_pos), continue; end
            R = norm(entry.true_pos);
            dist_all(end+1) = R/1e3;
            detect_all(end+1) = isfield(entry,'track_pos') && ~isempty(entry.track_pos) && ...
                                norm(entry.track_pos - entry.true_pos) < 1000;
            id_all_stat(end+1) = entry.id;
        end
        edges = 30:20:250; centers = (edges(1:end-1)+edges(2:end))/2;
        Pd_sim = nan(3, length(centers));
        for k = 1:3
            for b = 1:length(centers)
                idx = id_all_stat==k & dist_all>=edges(b) & dist_all<edges(b+1);
                if sum(idx) >= 5
                    Pd_sim(k,b) = mean(detect_all(idx));
                end
            end
        end
    end

    figure('Name','检测概率曲线','Color','w','Position',[100,100,800,600]);
    for k = 1:3
        subplot(1,3,k); hold on; grid on;
        plot(dist_grid/1e3, Pd_theory_single(k,:), '--', 'Color', colors{k}, 'LineWidth',1.5);
        plot(dist_grid/1e3, Pd_theory_multi(k,:), '-', 'Color', colors{k}, 'LineWidth',2);
        if exist('Pd_sim','var')
            plot(centers, Pd_sim(k,:), 'o', 'Color', colors{k}, 'MarkerFaceColor', colors{k});
        end
        xlabel('距离 (km)'); ylabel('P_d'); ylim([0 1]);
        title([target_names{k}, ' (RCS=',num2str(RCS_list(k)),' m^2)']);
        legend('单链路理论','多链路理论','仿真统计','Location','southwest');
    end
    sgtitle('图1 检测概率 vs 距离');
    Pd_theory = Pd_theory_multi;
catch ME
    warning('图1生成失败: %s', ME.message);
end

% ---------- 图2：雷达探测威力（理论） ----------
try
    figure('Name','雷达探测威力','Color','w');
    RCS_dB = linspace(-30, 20, 200);
    RCS_lin = 10.^(RCS_dB/10);
    Pd_target = 0.9;
    SNR_req = albersheim_snr_db_fixed(Pfa, Pd_target, 1);
    R_max = R0_m * 10.^((SNR0_dB - SNR_req + 10*log10(RCS_lin))/40);
    plot(RCS_dB, R_max/1e3, 'b-', 'LineWidth',1.5);
    grid on; xlabel('RCS (dBsm)'); ylabel('最大探测距离 (km)');
    title(['图2 雷达探测威力 (P_d=',num2str(Pd_target),', P_{fa}=',num2str(Pfa),')']);
catch ME
    warning('图2生成失败: %s', ME.message);
end

% ---------- 图3：定位精度理论CRLB vs 距离 ----------
try
    figure('Name','定位精度理论CRLB vs 距离','Color','w');
    dist_grid_crlb = linspace(30, 250, 50) * 1e3;
    crlb_grid = zeros(3, length(dist_grid_crlb));
    for k = 1:3
        RCS = RCS_list(k);
        for j = 1:length(dist_grid_crlb)
            xt = [dist_grid_crlb(j); 0; 10000];
            vt = [0;0;0];
            crlb_grid(k,j) = compute_crlb_signal(xt, vt, RCS, 1:L);
        end
        semilogy(dist_grid_crlb/1e3, crlb_grid(k,:), 'Color', colors{k}, 'LineWidth',1.5); hold on;
    end
    grid on; xlabel('距离 (km)'); ylabel('CRLB (m)');
    title('图3 定位精度理论下界 vs 距离');
    legend(target_names, 'Location','best');
    ylim([1 1e6]);
catch ME
    warning('图3生成失败: %s', ME.message);
end

% ---------- 图4：跟踪精度（实际误差 vs 时间） ----------
try
    figure('Name','跟踪精度（实际误差 vs 时间）','Color','w');
    plot(t_rt, e1, 'r-', t_rt, e2, 'g-', t_rt, e3, 'b-', 'LineWidth',1.2);
    grid on; xlabel('时间 (s)'); ylabel('位置误差 (m)');
    legend(target_names); title('图4 跟踪精度（实际误差）'); ylim([0 1000]);
catch ME
    warning('图4生成失败: %s', ME.message);
end

% ---------- 图5：多目标分辨能力 ----------
try
    res_range = sys.r_res;
    res_vel = lambda/(2*Tc);
    figure('Name','多目标分辨能力','Color','w');
    subplot(2,1,1); hold on; grid on;
    idx1 = find(id_all==1); idx2 = find(id_all==2); idx3 = find(id_all==3);
    if ~isempty(idx1) && ~isempty(idx2) && ~isempty(idx3)
        t1 = t_all(idx1); t2 = t_all(idx2); t3 = t_all(idx3);
        t_uniform = 0:0.5:60;
        x1 = interp1(t1, xt_all(1,idx1)', t_uniform, 'linear', NaN);
        x2 = interp1(t2, xt_all(1,idx2)', t_uniform, 'linear', NaN);
        x3 = interp1(t3, xt_all(1,idx3)', t_uniform, 'linear', NaN);
        y1 = interp1(t1, xt_all(2,idx1)', t_uniform, 'linear', NaN);
        y2 = interp1(t2, xt_all(2,idx2)', t_uniform, 'linear', NaN);
        y3 = interp1(t3, xt_all(2,idx3)', t_uniform, 'linear', NaN);
        z1 = interp1(t1, xt_all(3,idx1)', t_uniform, 'linear', NaN);
        z2 = interp1(t2, xt_all(3,idx2)', t_uniform, 'linear', NaN);
        z3 = interp1(t3, xt_all(3,idx3)', t_uniform, 'linear', NaN);
        d12 = sqrt((x1-x2).^2 + (y1-y2).^2 + (z1-z2).^2);
        d13 = sqrt((x1-x3).^2 + (y1-y3).^2 + (z1-z3).^2);
        d23 = sqrt((x2-x3).^2 + (y2-y3).^2 + (z2-z3).^2);
        plot(t_uniform, d12, 'r-', 'DisplayName','Jet1-Jet2');
        plot(t_uniform, d13, 'g-', 'DisplayName','Jet1-Jet3');
        plot(t_uniform, d23, 'b-', 'DisplayName','Jet2-Jet3');
        yline(res_range, 'k--', '距离分辨率');
        ylabel('距离差 (m)'); xlabel('时间 (s)'); legend show;
        title('目标间距离差 vs 分辨率');
    else
        text(0.5,0.5,'目标轨迹数据不足','HorizontalAlignment','center');
    end

    subplot(2,1,2); hold on; grid on;
    if ~isempty(idx1) && ~isempty(idx2) && ~isempty(idx3)
        v1 = vecnorm(vt_all(:,idx1),2,1); v1 = interp1(t1, v1, t_uniform, 'linear', NaN);
        v2 = vecnorm(vt_all(:,idx2),2,1); v2 = interp1(t2, v2, t_uniform, 'linear', NaN);
        v3 = vecnorm(vt_all(:,idx3),2,1); v3 = interp1(t3, v3, t_uniform, 'linear', NaN);
        dv12 = abs(v1 - v2);
        dv13 = abs(v1 - v3);
        dv23 = abs(v2 - v3);
        plot(t_uniform, dv12, 'r-', 'DisplayName','Jet1-Jet2');
        plot(t_uniform, dv13, 'g-', 'DisplayName','Jet1-Jet3');
        plot(t_uniform, dv23, 'b-', 'DisplayName','Jet2-Jet3');
        yline(res_vel, 'k--', '速度分辨率');
        ylabel('速度差 (m/s)'); xlabel('时间 (s)'); legend show;
        title('目标间速度差 vs 分辨率');
    else
        text(0.5,0.5,'目标轨迹数据不足','HorizontalAlignment','center');
    end
    sgtitle('图5 多目标分辨能力');
catch ME
    warning('图5生成失败: %s', ME.message);
end

% ---------- 图6：鬼影率（基于实际检测点统计）----------
try
    fprintf('正在统计鬼影率...\n');
    total_dets = 0;
    ghost_dets = 0;
    for frame = 1:length(all_dets_history)
        dets_frame = all_dets_history{frame};
        if isempty(dets_frame), continue; end
        t_frame = dets_frame(1).time;
        idx_t = find(abs(t_all - t_frame) < 0.01);
        true_targets = [];
        for i = 1:length(idx_t)
            tid = id_all(idx_t(i));
            true_targets = [true_targets; struct('pos', xt_all(:,idx_t(i))', 'vel', vt_all(:,idx_t(i))', 'id', tid)];
        end
        for d = 1:length(dets_frame)
            total_dets = total_dets + 1;
            is_true = false;
            link_id = dets_frame(d).link_id;
            if link_id < 1 || link_id > 6, continue; end
            [tx_idx, rx_idx] = ind2sub([2,3], link_id);
            if tx_idx<1 || tx_idx>2 || rx_idx<1 || rx_idx>3, continue; end
            tx = tx_pos(tx_idx,:);
            rx = rx_pos(rx_idx,:);
            for tgt = 1:length(true_targets)
                p = true_targets(tgt).pos;
                v = true_targets(tgt).vel;
                R_tx = norm(p - tx);
                R_rx = norm(p - rx);
                r_true = 0.5 * (R_tx + R_rx);
                u_tx = (p - tx)/R_tx;
                u_rx = (p - rx)/R_rx;
                v_true = 0.5 * (dot(v, u_tx) + dot(v, u_rx));
                if abs(dets_frame(d).r_meas - r_true) < 500 && abs(dets_frame(d).v_meas - v_true) < 50
                    is_true = true;
                    break;
                end
            end
            if ~is_true
                ghost_dets = ghost_dets + 1;
            end
        end
    end
    if total_dets > 0
        ghost_rate = ghost_dets / total_dets;
    else
        ghost_rate = 0;
        warning('没有检测点，鬼影率设为0。');
    end
    figure('Name','鬼影率统计','Color','w');
    bar(ghost_rate);
    set(gca, 'XTickLabel', {'鬼影率'});
    ylabel('比例'); title(sprintf('图6 实际鬼影率 = %.2f%% (基于%d个检测点)', ghost_rate*100, total_dets));
    grid on;
catch ME
    warning('图6生成失败: %s', ME.message);
end

% ---------- 图7：搜索效率（累积检测概率 vs 距离）----------
try
    dist_km = dist_grid/1e3;
    Pd_avg = mean(Pd_theory, 1);
    SE_cum = cumtrapz(dist_km, Pd_avg);
    figure('Name','搜索效率','Color','w');
    plot(dist_km, SE_cum, 'b-', 'LineWidth',1.5);
    grid on; xlabel('距离 (km)'); ylabel('累积积分值 (km)');
    title('图7 搜索效率（检测概率累积积分）');
    text(0.5*max(dist_km), 0.5*max(SE_cum), sprintf('总效率 = %.1f km', SE_cum(end)), ...
        'HorizontalAlignment','center');
catch ME
    warning('图7生成失败: %s', ME.message);
end

% ---------- 图8：最大稳定跟踪距离 ----------
try
    figure('Name','最大稳定跟踪距离','Color','w');
    for k = 1:3
        switch k
            case 1, err = e1; case 2, err = e2; case 3, err = e3;
        end
        idx_fail = find(err > 1000, 1);
        if ~isempty(idx_fail)
            t_fail = t_rt(idx_fail);
            [~, idx_t] = min(abs(t_all - t_fail));
            R_fail = norm(xt_all(:,idx_t));
            bar(k, R_fail/1e3, 'FaceColor', colors{k});
            text(k, R_fail/1e3+5, sprintf('%.0f km', R_fail/1e3), 'HorizontalAlignment','center');
        else
            idx_target = find(id_all == k);
            if ~isempty(idx_target)
                R_max_target = max(vecnorm(xt_all(:,idx_target),2,1));
                bar(k, R_max_target/1e3, 'FaceColor', colors{k});
                text(k, R_max_target/1e3+5, sprintf('>%.0f km*', R_max_target/1e3), 'HorizontalAlignment','center');
            else
                bar(k, 0, 'FaceColor', colors{k});
                text(k, 5, '无数据', 'HorizontalAlignment','center');
            end
        end
        hold on;
    end
    grid on; xlabel('目标'); ylabel('最大稳定跟踪距离 (km)');
    set(gca, 'XTick', 1:3, 'XTickLabel', target_names);
    title('图8 误差首次超过1000m时的距离（*表示从未超过）');
catch ME
    warning('图8生成失败: %s', ME.message);
end

% ---------- 图9：信号级融合优势 ----------
try
    figure('Name','信号级融合优势','Color','w');
    for k = 1:3
        subplot(1,3,k); hold on; grid on;
        idx = find(id_all == k);
        if isempty(idx)
            text(0.5,0.5,'无数据','HorizontalAlignment','center');
            title(target_names{k});
            continue;
        end
        t_k = t_all(idx);  % 列向量
        crlb_sig = zeros(length(idx),1);
        crlb_info = zeros(length(idx),1);
        for i = 1:length(idx)
            xt = xt_all(:,idx(i));
            vt = vt_all(:,idx(i));
            RCS = RCS_list(k);
            crlb_sig(i) = compute_crlb_signal(xt, vt, RCS, 1:L);
            crlb_info(i) = compute_crlb_info_range(xt, RCS, 1:L);
        end
        % 过滤无效值
        valid_sig = ~isnan(crlb_sig);
        valid_info = ~isnan(crlb_info);
        if any(valid_sig)
            plot(t_k(valid_sig), crlb_sig(valid_sig), 'b-', 'LineWidth',1.5, 'DisplayName','信号级CRLB');
        end
        if any(valid_info)
            plot(t_k(valid_info), crlb_info(valid_info), 'r--', 'LineWidth',1.5, 'DisplayName','信息级CRLB');
        end
        switch k
            case 1, err = e1; case 2, err = e2; case 3, err = e3;
        end
        plot(t_rt, err, 'k.', 'MarkerSize',4, 'DisplayName','实际误差');
        xlabel('时间 (s)'); ylabel('位置RMSE (m)');
        title(target_names{k}); legend show; ylim([0 1000]);
    end
    sgtitle('图9 信号级 vs 信息级融合下界与实际误差');
catch ME
    warning('图9生成失败: %s', ME.message);
end

% ---------- 图10：距离多普勒融合（生成理想RD图，避免数据时序问题）----------
try
    t_plot = 30;
    [~, tidx] = min(abs(t_all - t_plot));
    if ~isempty(tidx)
        xt_plot = xt_all(:,tidx);
        vt_plot = vt_all(:,tidx);
        figure('Name','距离多普勒融合','Color','w');
        % 计算目标所在码元（以第一个链路为基准）
        d_tx = norm(xt_plot - tx_pos(1,:));
        d_rx = norm(xt_plot - rx_pos(1,:));
        r_eq = 0.5 * (d_tx + d_rx);
        cid_target = floor(r_eq / sys.R_unamb) + 1;
        cid_target = max(1, min(cid_target, sys.Ncodes));
        % 构造该码元的距离轴
        dist_start = (cid_target-1) * sys.Nsc * sys.r_res / 1e3;
        dist_end = cid_target * sys.Nsc * sys.r_res / 1e3;
        x_axis = linspace(dist_start, dist_end, sys.Nsc);
        % 生成每个链路的理想RD图
        for link = 1:L
            [i,j] = ind2sub([2,3], link);
            tx = tx_pos(i,:);
            rx = rx_pos(j,:);
            RD_link = generate_ideal_rd(tx, rx, xt_plot, vt_plot, sys); % 返回 Nsc × Nc
            subplot(2,3,link);
            % 注意：imagesc(x,y,C) 要求 x 长度为 size(C,2)，y 长度为 size(C,1)
            % RD_link 行数 = Nsc (距离)，列数 = Nc (速度)
            % 因此横轴用 x_axis (距离)，纵轴用 sys.v_axis (速度)
            imagesc(x_axis, sys.v_axis, 20*log10(abs(RD_link)+1e-12).');
            axis xy; xlabel('距离 (km)'); ylabel('速度 (m/s)');
            title(sprintf('链路 %d: Tx%d-Rx%d', link, i, j));
            colormap jet; colorbar;
        end
        % 融合RD图（能量叠加）
        RD_fused = zeros(sys.Nsc, Nc);
        for link = 1:L
            [i,j] = ind2sub([2,3], link);
            tx = tx_pos(i,:);
            rx = rx_pos(j,:);
            RD_link = generate_ideal_rd(tx, rx, xt_plot, vt_plot, sys);
            RD_fused = RD_fused + abs(RD_link).^2;
        end
        subplot(2,3,[4,5,6]);
        imagesc(x_axis, sys.v_axis, 20*log10(sqrt(RD_fused)+1e-12).');
        axis xy; xlabel('距离 (km)'); ylabel('速度 (m/s)');
        title('融合RD图（能量叠加）'); colorbar;
        sgtitle('图10 各链路RD图与融合RD图 (t=30s)');
    else
        warning('图10: 时刻t=30无目标轨迹，跳过RD图。');
    end
catch ME
    warning('图10生成失败: %s', ME.message);
end

% ---------- 图11：三通道MTI频率响应 ----------
try
    figure('Name','三通道MTI频率响应','Color','w');
    f = linspace(-1/(2*PRI), 1/(2*PRI), 1000);
    H1 = freqz([1,-1],1,f,PRI,'whole');
    H2 = freqz([1,-2,1],1,f,PRI,'whole');
    H3 = freqz([1,-3,3,-1],1,f,PRI,'whole');
    plot(f, 20*log10(abs(H1)), 'b-', 'DisplayName','1阶MTI (SLOW)'); hold on;
    plot(f, 20*log10(abs(H2)), 'g-', 'DisplayName','2阶MTI (FAST)');
    plot(f, 20*log10(abs(H3)), 'r-', 'DisplayName','3阶MTI (VFAST)');
    xline(0, 'k--'); xline(1/(2*PRI), 'k--');
    xlabel('频率 (Hz)'); ylabel('幅度响应 (dB)'); grid on; legend show;
    v2fd = @(v) 2*v/lambda;
    fd1 = v2fd(mean([targets(1).vel])); fd2 = v2fd(mean([targets(2).vel])); fd3 = v2fd(mean([targets(3).vel]));
    xline(fd1, '--r', 'Color', 'r', 'DisplayName', 'Jet1'); 
    xline(fd2, '--g', 'Color', 'g', 'DisplayName', 'Jet2');
    xline(fd3, '--b', 'Color', 'b', 'DisplayName', 'Jet3');
    title('图11 三通道MTI频率响应与目标速度标注');
catch ME
    warning('图11生成失败: %s', ME.message);
end

% ---------- 图12：目标散射特性（不同RCS对定位精度的影响）----------
try
    figure('Name','目标散射特性','Color','w');
    dist_grid_rcs = linspace(30, 250, 50) * 1e3;
    % 固定目标位置（例如在x轴上，高度10km）
    xt_fixed = [100e3; 0; 10000];  % 临时占位，后面会修改距离
    vt_fixed = [0;0;0];
    rcs_values = [25, 5, 0.5];
    rcs_names = {'RCS=25 m²','RCS=5 m²','RCS=0.5 m²'};
    crlb_rcs = zeros(length(rcs_values), length(dist_grid_rcs));
    for r = 1:length(rcs_values)
        RCS = rcs_values(r);
        for j = 1:length(dist_grid_rcs)
            xt_fixed(1) = dist_grid_rcs(j); % 改变距离
            crlb_rcs(r,j) = compute_crlb_signal(xt_fixed, vt_fixed, RCS, 1:L);
        end
        semilogy(dist_grid_rcs/1e3, crlb_rcs(r,:), 'LineWidth',1.5); hold on;
    end
    grid on; xlabel('距离 (km)'); ylabel('CRLB (m)');
    title('图12 目标散射特性对定位精度的影响');
    legend(rcs_names, 'Location','best');
    ylim([1 1e6]);
catch ME
    warning('图12生成失败: %s', ME.message);
end

% ---------- 图13：节点数量与位置影响 ----------
try
    % 节点数量影响
    figure('Name','节点数量对定位精度的影响','Color','w');
    t_fixed = 30;
    [~, tidx] = min(abs(t_all - t_fixed));
    if ~isempty(tidx)
        xt = xt_all(:,tidx);
        vt = vt_all(:,tidx);
        for k = 1:3
            RCS = RCS_list(k);
            crlb_vs_n = zeros(1,L);
            for n = 1:L
                crlb_vs_n(n) = compute_crlb_signal(xt, vt, RCS, 1:n);
            end
            plot(1:L, crlb_vs_n, 'o-', 'Color', colors{k}, 'LineWidth',1.5); hold on;
        end
        grid on; xlabel('使用的链路数'); ylabel('CRLB (m)');
        title(['图13a 节点数量对CRLB的影响 (t=',num2str(t_fixed),'s)']);
        legend(target_names, 'Location','best');
    else
        text(0.5,0.5,'无t=30s轨迹数据','HorizontalAlignment','center');
        title('图13a 无数据');
    end

    % 节点位置影响（修正布局定义，确保接收机数量为3）
    layouts = {
        '集中式',  repmat(mean(rx_pos,1), 3, 1);  % 三个接收机集中在质心
        '分布式',  rx_pos;                        % 原分布式布局
        '一字型',  [50000,0,5000; 50000,10000,5000; 50000,-10000,5000]  % 一字型布局（仅接收机）
    };
    figure('Name','节点位置对定位精度的影响','Color','w');
    for layout_idx = 1:size(layouts,1)
        rx_layout = layouts{layout_idx,2}; % 直接是3x3矩阵
        avg_crlb = zeros(1,3);
        for k = 1:3
            idx = find(id_all == k);
            if isempty(idx)
                avg_crlb(k) = NaN;
                continue;
            end
            crlb_sum = 0; cnt = 0;
            for i = 1:length(idx)
                xt = xt_all(:,idx(i));
                vt = vt_all(:,idx(i));
                RCS = RCS_list(k);
                rx_pos_orig = rx_pos;
                rx_pos = rx_layout;
                crlb_val = compute_crlb_signal(xt, vt, RCS, 1:L);
                rx_pos = rx_pos_orig;
                if ~isnan(crlb_val)
                    crlb_sum = crlb_sum + crlb_val;
                    cnt = cnt + 1;
                end
            end
            if cnt>0
                avg_crlb(k) = crlb_sum / cnt;
            else
                avg_crlb(k) = NaN;
            end
        end
        subplot(2,2,layout_idx);
        bar(avg_crlb, 'FaceColor', [0.7 0.7 0.7]);
        set(gca, 'XTickLabel', target_names);
        ylabel('平均CRLB (m)'); title(layouts{layout_idx,1});
        grid on;
    end
    sgtitle('图13b 不同节点布局下的平均定位精度');
catch ME
    warning('图13生成失败: %s', ME.message);
end

% ---------- 图14：信号带宽的影响 ----------
try
    figure('Name','信号带宽对定位精度的影响','Color','w');
    B_list = [0.5, 1, 2, 4, 8] * 1e6;
    t_fixed = 30;
    [~, tidx] = min(abs(t_all - t_fixed));
    if ~isempty(tidx)
        xt = xt_all(:,tidx);
        vt = vt_all(:,tidx);
        for k = 1:3
            RCS = RCS_list(k);
            crlb_vs_B = zeros(size(B_list));
            for b = 1:length(B_list)
                B_orig = B;
                B = B_list(b);
                crlb_vs_B(b) = compute_crlb_signal(xt, vt, RCS, 1:L);
                B = B_orig;
            end
            semilogx(B_list/1e6, crlb_vs_B, 'o-', 'Color', colors{k}, 'LineWidth',1.5); hold on;
        end
        grid on; xlabel('带宽 (MHz)'); ylabel('CRLB (m)');
        title(['图14 带宽对CRLB的影响 (t=',num2str(t_fixed),'s)']);
        legend(target_names, 'Location','best');
    else
        text(0.5,0.5,'无t=30s轨迹数据','HorizontalAlignment','center');
        title('图14 无数据');
    end
catch ME
    warning('图14生成失败: %s', ME.message);
end

% ---------- 图15：系统误差的影响 ----------
try
    figure('Name','系统误差的影响','Color','w');
    % 子图1：相位误差对相干增益的影响
    subplot(1,2,1);
    sigma_phi_deg = linspace(0, 90, 50);
    sigma_phi = deg2rad(sigma_phi_deg);
    L_avail = 6;
    G_coh = 10*log10( (1 + (L_avail-1)*exp(-sigma_phi.^2)) / L_avail );
    plot(sigma_phi_deg, G_coh, 'b-', 'LineWidth',1.5);
    grid on; xlabel('相位误差标准差 (度)'); ylabel('相干增益 (dB)');
    title('相位误差对相干合成增益的影响');
    yline(10*log10(L_avail), 'k--', '理论最大');

    % 子图2：站址误差对CRLB的影响（蒙特卡洛平均）
    subplot(1,2,2);
    sigma_site = linspace(0, 50, 20);
    t_fixed = 30;
    [~, tidx] = min(abs(t_all - t_fixed));
    if ~isempty(tidx)
        xt = xt_all(:,tidx);
        vt = vt_all(:,tidx);
        for k = 1:3
            RCS = RCS_list(k);
            crlb_vs_site = zeros(size(sigma_site));
            for s = 1:length(sigma_site)
                n_mc = 20;
                crlb_sum = 0; cnt_mc = 0;
                for mc = 1:n_mc
                    tx_err = tx_pos + sigma_site(s)*randn(size(tx_pos));
                    rx_err = rx_pos + sigma_site(s)*randn(size(rx_pos));
                    tx_pos_orig = tx_pos; rx_pos_orig = rx_pos;
                    tx_pos = tx_err; rx_pos = rx_err;
                    crlb_val = compute_crlb_signal(xt, vt, RCS, 1:L);
                    tx_pos = tx_pos_orig; rx_pos = rx_pos_orig;
                    if ~isnan(crlb_val)
                        crlb_sum = crlb_sum + crlb_val;
                        cnt_mc = cnt_mc + 1;
                    end
                end
                if cnt_mc > 0
                    crlb_vs_site(s) = crlb_sum / cnt_mc;
                else
                    crlb_vs_site(s) = NaN;
                end
            end
            plot(sigma_site, crlb_vs_site, 'o-', 'Color', colors{k}, 'LineWidth',1.5); hold on;
        end
        grid on; xlabel('站址误差标准差 (m)'); ylabel('CRLB (m)');
        title('站址误差对CRLB的影响 (蒙特卡洛平均)');
        legend(target_names, 'Location','best');
    else
        text(0.5,0.5,'无t=30s轨迹数据','HorizontalAlignment','center');
        title('无数据');
    end
    sgtitle('图15 系统误差对多节点协同效能的影响');
catch ME
    warning('图15生成失败: %s', ME.message);
end

end