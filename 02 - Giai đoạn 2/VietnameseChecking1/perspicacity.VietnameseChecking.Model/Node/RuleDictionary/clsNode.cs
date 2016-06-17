// Dự án:   BẮT LỖI CHÍNH TẢ TIẾNG VIỆT
// Tác giả: Lê Viết Mẫn
// Email:   manleviet@yahoo.com.hk

using System;

namespace perspicacity.VietnameseChecking.Model.Node.RuleDictionary
{
	/// <summary>
	/// Lớp _node này biểu diễn cho một nút trong cấu trúc
	/// từ điển Rule.
	/// </summary>
	public class _node : _base
	{
		#region Vùng hàm khởi tạo
		/// <summary>
		/// Cấu tử cơ bản, không có tham số.
		/// </summary>
		public _node():base(){}

		/// <summary>
		/// Cấu tử thứ hai, có hai tham số.
		/// </summary>
		/// <param name="leftrule">
		/// Nhận vế phải của luật sinh.
		/// </param>
		/// <param name="rightrule">
		/// Nhận vế trái của luật sinh.
		/// </param>
		public _node(string leftrule, string rightrule) : base(leftrule,rightrule){}
		#endregion

		#region Vùng thuộc tính
		/// <summary>
		/// Thuộc tính LeftRule, lấy hay gán vế phải của luật sinh.
		/// </summary>
		public string LeftRule
		{
			get
			{
				return base.Content1;
			}
			set
			{
				base.Content1 = value;
			}
		}

		/// <summary>
		/// Thuộc tính RightRule, lấy hay gán vế trái của luật sinh.
		/// </summary>
		public string RightRule
		{
			get
			{
				return base.Content2;
			}
			set
			{
				base.Content2 = value;
			}
		}
		#endregion
	}
}