using System.Collections.Generic;

namespace overlay_gpt
{
    public interface IContextWriter
    {
        /// <summary>
        /// 백그라운드 프로세스 여부를 나타냅니다.
        /// </summary>
        bool IsTargetProg { get; set; }

        /// <summary>
        /// 파일을 백그라운드에서 엽니다.
        /// </summary>
        /// <param name="filePath">열 파일의 경로</param>
        /// <returns>파일 열기 성공 여부</returns>
        bool OpenFile(string filePath);

        /// <summary>
        /// HTML 형식의 텍스트를 파일에 적용합니다.
        /// </summary>
        /// <param name="text">HTML 형식의 텍스트</param>
        /// <param name="lineNumber">적용할 위치 정보</param>
        /// <returns>작업 성공 여부</returns>
        bool ApplyTextWithStyle(string text, string lineNumber);

        /// <summary>
        /// 파일 정보를 가져옵니다.
        /// </summary>
        /// <returns>파일 ID, 볼륨 ID, 파일 타입, 파일 이름, 파일 경로</returns>
        (ulong? FileId, uint? VolumeId, string FileType, string FileName, string FilePath) GetFileInfo();
    }
}