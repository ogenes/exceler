<?php


namespace Ogenes\Exceler;

use RuntimeException;

/**
 * User: ogenes <ogenes.yi@gmail.com>
 * Date: 2022/7/23
 */
class DownloadClient
{
    private static $instance;
    
    public static function getInstance(): self
    {
        self::$instance instanceof self || self::$instance = new self();
        return self::$instance;
    }
    
    private $imageType = 'jpeg';//图片格式
    
    private const ALLOW_TYPE = [
        'jpeg',
        'png',
        'Jpeg',
        'PNG'
    ];
    
    
    public function multiDownloadImg(array $requests, string $dir, array $options = [])
    {
        if (empty($requests)) {
            return null;
        }
        if (!is_dir($dir) && !mkdir($dir, 0700, true) && !is_dir($dir)) {
            throw new RuntimeException(sprintf('Directory "%s" was not created', $dir));
        }
        
        $queue = curl_multi_init();
        $map = [];
        foreach ($requests as $id => $request) {
            $ch = curl_init();
            curl_setopt($ch, CURLOPT_URL, $request['url']);
            $timeout = !empty($options['TIMEOUT']) ? (int)$options['TIMEOUT'] : 30;
            curl_setopt($ch, CURLOPT_TIMEOUT, $timeout);
            curl_setopt($ch, CURLOPT_RETURNTRANSFER, 1);
            curl_setopt($ch, CURLOPT_FOLLOWLOCATION, 1);
            curl_setopt($ch, CURLOPT_HEADER, true);
            curl_multi_add_handle($queue, $ch);
            $map[(string)$ch] = $id;
        }
        $responses = [];
        do {
            do {
                $code = curl_multi_exec($queue, $active);
            } while ($code === CURLM_CALL_MULTI_PERFORM);
            
            if ($code !== CURLM_OK) {
                break;
            }
            while ($done = curl_multi_info_read($queue)) {
                $data = curl_multi_getcontent($done['handle']);
                $headerSize = curl_getinfo($done['handle'], CURLINFO_HEADER_SIZE);
                $header = substr($data, 0, $headerSize);
                $file_content = substr($data, $headerSize);
                $id = $map[(string)$done['handle']];
//                $retCode = curl_errno($done['handle']);
//                if ($retCode > 0) {
//                    $responses[$id] = '';
//                    continue;
//                }
                curl_multi_remove_handle($queue, $done['handle']);
                curl_close($done['handle']);
                
                $pattern = "/[c|C]ontent-[t|T]ype: images?\/(?<type>(?:" . implode('|', self::ALLOW_TYPE) . "))/";
                preg_match($pattern, $header, $matches);
                
                if (isset($matches['type']) && in_array($matches['type'], self::ALLOW_TYPE, false)) {
                    $this->imageType = $matches['type'];
                } else {
                    
                    $responses[$id] = '';
                    continue;
                }
                $filename = $dir . $id . '.' . $this->imageType;
                $fp = fopen($filename, 'w');
                fwrite($fp, $file_content);
                fclose($fp);
                $responses[$id] = $filename;
            }
            
            if ($active > 0) {
                curl_multi_select($queue, 0.5);
            }
        } while ($active);
        
        curl_multi_close($queue);
        return $responses;
    }
    
    
    /**
     * 断点续传
     * @param $file
     * @param $file_display_name
     * @param int $chunk
     */
    public function breaking($file, $file_display_name, int $chunk = 2048): void
    {
        if (!file_exists($file)) {
            echo '文件不存在！(不支持中文文件名)';
            exit;
        }
        $fSize = @filesize($file);
        if (!empty($fSize)) {
            $start = null;
            $end = $fSize - 1;
            if (isset($_SERVER['HTTP_RANGE']) && ($_SERVER['HTTP_RANGE'] !== "") && preg_match("/^bytes=([0-9]+)-([0-9]*)$/i", $_SERVER['HTTP_RANGE'], $match) && ($match[1] < $fSize) && ($match[2] < $fSize)) {
                $start = $match[1];
                if (!empty($match[2])) {
                    $end = $match[2];
                }
            }
            header('Cache-control: public');
            header('Pragma: public');
            if ($start === null) {
                header('HTTP/1.1 200 OK');
                header("Content-Length: $fSize");
                header('Accept-Ranges: bytes');
            } else {
                header('HTTP/1.1 206 Partial Content');
                header('Content-Length: ' . ($end - $start + 1));
                header('Content-Ranges: bytes ' . $start . '-' . $end . '/' . $fSize);
            }
            header('Content-Type: application/octet-stream');
            header('Content-Disposition: attachment; filename=' . $file_display_name);
            ob_clean();
            flush();
            $fp = fopen($file, 'rb');
            fseek($fp, $start);
            while (($nowNum = ftell($fp)) < $end) {
                if ($nowNum >= ($end - $chunk)) {
                    $chunk = $end - $nowNum + 1;
                }
                echo fread($fp, $chunk);
            }
            fclose($fp);
        }
    }
    
    /**
     * 直接下载
     *
     * @param $file
     * @param $file_display_name
     */
    public function direct($file, $file_display_name): void
    {
        if (!file_exists($file)) {
            echo '文件不存在！(不支持中文文件名)';
            exit;
        }
        $fSize = @filesize($file);
        $fp = fopen($file, 'rb');
        header("Content-Type: application/octet-stream");
        header("Accept-Ranges: bytes");
        header("Accept-Length: " . $fSize);
        header("Content-Disposition: attachment; filename=$file_display_name");
        echo fread($fp, $fSize);
        fclose($fp);
    }
    
    /**
     * 远程文件下载
     * @param $url
     * @param $filepath
     * @return string
     * User: Ogenes
     * Date: 2023/9/16
     */
    public function downloadFile($url, $filepath): string
    {
        if (empty($url) || empty($filepath)) {
            return '';
        }
        $fp = fopen($filepath, 'w');
        $ch = curl_init($url);
        curl_setopt($ch, CURLOPT_FILE, $fp);
        curl_exec($ch);
        curl_close($ch);
        return $filepath;
    }
}


