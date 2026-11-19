// Cấu hình dự án mặc định (Coastal). Dự án khác = bản ghi cấu hình khác, không sửa code.
import type { ProjectConfig } from './types';

export const RFI_OPTIONS_COASTAL = [
  { code: '1', text: 'Xem trả lời bên dưới', result: 'CLOSE' as const },
  { code: '2', text: 'Vui lòng cung cấp thêm thông tin', result: 'NEED_INFO' as const },
  { code: '3', text: 'Chuyển tiếp tới', result: 'FORWARD' as const },
];
export const RFA_OPTIONS_COASTAL = [
  { code: 'A', text: 'Chấp thuận yêu cầu', result: 'CLOSE' as const },
  { code: 'B', text: 'Chấp thuận yêu cầu với điều kiện hoặc ý kiến phản hồi bên dưới', result: 'CLOSE' as const },
  { code: 'C', text: 'Vui lòng xem ý kiến phản hồi và đệ trình lại', result: 'RESUBMIT' as const },
  { code: 'D', text: 'Từ chối chấp thuận yêu cầu', result: 'REJECT' as const },
];

export function coastalConfig(members: ProjectConfig['members'] = {}): ProjectConfig {
  const p = '{yymmdd}-{project}-{contractor}-{type}-{seq:3}';
  return {
    projectId: 'COASTAL_QNG', code: 'COA', contractor: 'TBC', timezone: 'Asia/Ho_Chi_Minh', configVersion: 1, members,
    types: {
      RFI: { pattern: p, slaDays: 3, reviewer: 'CĐT', options: RFI_OPTIONS_COASTAL },
      RFA: { pattern: p, slaDays: 1, reviewer: 'CĐT', options: RFA_OPTIONS_COASTAL },
      RFC: { pattern: p, slaDays: 3, reviewer: 'CĐT', options: RFA_OPTIONS_COASTAL },
    },
  };
}
